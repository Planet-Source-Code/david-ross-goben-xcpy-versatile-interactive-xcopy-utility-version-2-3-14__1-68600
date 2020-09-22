Attribute VB_Name = "modXcopy"
Option Explicit
'~modXcopy.bas;modCheckDirs.bas;modExpandEnvStrings.bas;modGetFileAndDirData.bas;modAddSlash.bas;modNetPathFunctions.bas;
'Perform Xcopy function with numerous options

'copyfile for non-nt/xp
Public Declare Function CopyFileA Lib "kernel32" _
              (ByVal lpExistingFileName As String, _
               ByVal lpNewFileName As String, _
               ByVal bFailIfExists As Long) As Long

'copyfile for nt/xp
Public Declare Function CopyFileEx Lib "kernel32" Alias "CopyFileExA" _
              (ByVal lpExistingFileName As String, _
               ByVal lpNewFileName As String, _
               ByVal lpProgressRoutine As Long, _
               ByVal lpData As Long, _
               ByRef pbCancel As Long, _
               ByVal dwCopyFlags As Long) As Long

'return codes from the CopyFileEx callback
Public Const PROGRESS_CONTINUE As Long = 0
'Public Const PROGRESS_CANCEL As Long = 1
'Public Const PROGRESS_STOP As Long = 2
'Public Const PROGRESS_QUIET As Long = 3

'CopyFileEx callback state changes
Public Const CALLBACK_CHUNK_FINISHED As Long = &H0
Public Const CALLBACK_STREAM_SWITCH As Long = &H1

'CopyFileEx option flags
Public Const COPY_FILE_FAIL_IF_EXISTS As Long = &H1
Public Const COPY_FILE_RESTARTABLE As Long = &H2
Public Const COPY_FILE_OPEN_SOURCE_FOR_WRITE As Long = &H4

'variables devoted to the CopyFile functions
Public CopyCancel As Long         'set to 1 for canceling operation
Private CopyDirection As Boolean  'True = add dots, False = remove dots
Private CopyCntr As Long          'count to 10 (up to 10 dots)
Private CopyMask As Long          'count 32 before affecting CopyCtr
Private NT As Boolean             'True if NT/XP system

Private TempSrc As String     'temp source path
Private TempDst As String     'temp dest path

Private IncSubDir As Boolean  'include subdir flag
Private ModOnly As Boolean    'modified only flag
Private NewOnly As Boolean    'New only flag

Private ZipPath As String     'path to zip file
Private OrgZipName As String  'original zip filename
Private ZipName As String     'current zip file name
Private ZipStrip As Long      'index into Source path to strip off
Private ZipInc As Long        'zip file incrementer
Private LogHead As String     'header text for begin/end file process

Public Cancel As Boolean      'true when errors
Public LogFile As String      'log file
Public AppendLog As Boolean   'flag indicating log file should be appended to

Public StartTime As Date      'hold start time for time difference calcs
Public LastTime As String     'most current time checked
Public ElapsedTime As Date    'StartTime-LastTime
Public LogData As Collection  'store log entries
Public Pausing As Boolean     'pausing processing

Public fso As FileSystemObject 'file I/O resource

Private FolderCnt As Long     'number of folders found
Private FileCnt As Long       'number of files found
Private ItemCnt As Long       'number of items (folders and files)
Private MaxCnt As Long        'max count (found during counting)
Private Pcnt As Long          'percentage storage
Private OfFolders As String
Private OfFiles As String

'*******************************************************************************
' Function Name     : Xcopy
' Purpose           : Xcopy Src to Dst with options
'*******************************************************************************
Public Function Xcopy(SrcPath As String, _
                      DstPath As String, _
                      Optional IncludeSubDirs As Boolean = False, _
                      Optional ModifiedOnly As Boolean = False, _
                      Optional NewFilesOnly As Boolean = False, _
                      Optional CountData As Boolean = False) As Long
  Dim src As String, dst As String, S As String, Dta As String
  Dim srcFolder As Folder, dstFolder As Folder
  Dim srcFile As File, dstFile As File
  Dim Idx As Long
  Dim I As Integer, j As Integer, K As Integer
  Dim Ts As TextStream
  Dim Dt As Date
'
' if src is just drive, ensure a slash added
'
  NT = IsNT()         'set True if NT/XP
  If Right$(SrcPath, 1) = ":" Then SrcPath = AddSlash(SrcPath)  'if something like "D:"
  If Left$(SrcPath, 2) = "\\" Then
    I = InStr(3, SrcPath, "\")
    If CBool(I) Then
      I = InStr(I + 1, SrcPath, "\")
      If Not CBool(I) And Right$(SrcPath, 1) <> "\" Then SrcPath = AddSlash(SrcPath)
    End If
  End If
'
' if dest is just drive, ensure a slash added
'
  If Right$(DstPath, 1) = ":" Then DstPath = AddSlash(DstPath)  'if something like "D:"
  If Left$(DstPath, 2) = "\\" Then
    I = InStr(3, DstPath, "\")
    If CBool(I) Then
      I = InStr(I + 1, DstPath, "\")
      If Not CBool(I) And Right$(DstPath, 1) <> "\" Then DstPath = AddSlash(DstPath)
    End If
  End If
'
' init header for begin and end ession
'
  With frmXcopy
    S = " XCPY from " & SrcPath
    S = S & " to " & DstPath
    If .chkZip.Value = vbChecked And CBool(Len(.txtZIP.Text)) Then S = AddSlash(S) & Trim$(.txtZIP.Text)
    LogHead = S & " at: "
  End With
'
' now set up for processing
'
  StartTime = Now                                               'get start time
  LogDataAdd "Begin" & LogHead & CStr(StartTime)                'add to log
  With frmXcopy
    .lblStartTime.Caption = Format(StartTime, "hh:mm:ss")       'display in form
    LastTime = "00:00:00"                                       'null out end and interval times
    .lblElapsed.Caption = LastTime
    .lblEndTime.Caption = LastTime
'
' if we want to append to log file, flag it, and remove + from the start of the param if present
'
    If CBool(Len(LogFile)) Then
      If Left$(LogFile, 1) = "+" Then
        LogFile = Trim$(Mid$(LogFile, 2))
        AppendLog = True
      Else
        AppendLog = .chkAppend.Value = vbChecked
      End If
    Else
      AppendLog = False
    End If
  End With
'
' set other options
'
  IncSubDir = IncludeSubDirs              'set flag for including sub-directories
  ModOnly = ModifiedOnly                  'set flag for copying modified files only
  NewOnly = NewFilesOnly                  'set flag for copying new files only
'
' build solid source
'
  TempSrc = vbNullString                  'allow for tempdrive
  TempDst = vbNullString
  src = Trim$(SrcPath)                    'get source
  src = GetDriveFromNetPath(src)          'try to get a local connection
  If Right$(src, 1) = "\" And Len(src) = 3 Then src = Left$(src, Len(src) - 1)
  If Left$(src, 2) = "\\" Then TempSrc = GetFreeDriveLetter("M")
  If Len(TempSrc) Then                    'we've got a new drive
    If ConnectNetDrive(TempSrc, src) Then 'can we connect?
      src = GetDriveFromNetPath(src)      'yes, so reassign src to the new path
    Else
      TempSrc = vbNullString              'failed, so try to go on
    End If
  End If
'
' check for full wildcard specs. Keep up to last "\" if so
'
  If Right$(src, 4) = "\*.*" Then src = Left$(src, Len(src) - 3)
  If Right$(src, 2) = "\*" Then src = Left$(src, Len(src) - 1)
'
' check for invalid wildcard specs
'
  I = InStrRev(src, "\")                  'see if valid wildcards used, if any
  K = InStr(1, src, "*")
  j = InStr(1, src, "?")
  If (j > 0 And j < I) Or (K > 0 And K < I) Then  'wildcards and valid?
    Call ExitXcopy
    CenterMsgBoxOnForm frmXcopy, "Source Path cannot have wildcards below the last slash" & vbCrLf & _
           "Reference: " & SrcPath
    Exit Function
  End If
'
' strip trailing "\" if we had stripped full wildcards
'
  If Right$(SrcPath, 1) <> "\" Then src = RemoveSlash(src)
'
' check for a valid drive
'
  If Not fso.FolderExists(src) Then
    Call ExitXcopy
    CenterMsgBoxOnForm frmXcopy, "Source spec for " & SrcPath & " cannot be found"
    Exit Function
  End If
'
' build solid destination
'
  dst = Trim$(DstPath)                    'destination path
  dst = GetDriveFromNetPath(dst)          'see if we can get a local spec
  If Right$(dst, 1) = "\" And Len(dst) > 3 Then dst = Left$(dst, Len(dst) - 1)
  If Left$(dst, 2) = "\\" Then TempDst = GetFreeDriveLetter("M")
  If Len(TempDst) Then
    If ConnectNetDrive(TempDst, dst) Then
      src = GetDriveFromNetPath(dst)
    Else
      TempDst = vbNullString
    End If
  End If
'
' dest cannot have wildcards
'
  If CBool(InStr(1, dst, "*")) Or CBool(InStr(1, dst, "?")) Then
    LogDataAdd "Error: Destination Path cannot contain wildcards: " & DstPath
    Call ExitXcopy
    CenterMsgBoxOnForm frmXcopy, "Destination Path cannot contain wildcards" & vbCrLf & _
                                 "Reference: " & DstPath
    Exit Function
  End If
'
' if destination is a file
'
  If fso.FileExists(dst) Then
    LogDataAdd "Error: Destination path is not a directory: " & DstPath
    Call ExitXcopy
    CenterMsgBoxOnForm frmXcopy, "Destination path is not a directory: " & DstPath
    Exit Function
'
' check for valid destination path. Create if not found
'
  ElseIf Not fso.FolderExists(dst) Then
    CheckDirs dst                     'build path
    On Error Resume Next
    MkDir dst                         'make sure dst dir exists
    On Error GoTo 0
    If Not fso.FolderExists(dst) Then
      LogDataAdd "Error: Cannot create destination path: " & DstPath
      Call ExitXcopy
      CenterMsgBoxOnForm frmXcopy, "Cannot create destination path: " & DstPath
      Exit Function
    End If
  End If
'
' now check to see if we will be zipping the source files
'
  With frmXcopy
    If .chkZip.Value = vbChecked Then                         'user selected to zip them
      ZipName = Trim$(.txtZIP.Text)                           'get zip filename
      If CBool(Len(ZipName)) Then                             'if something there
        I = InStrRev(ZipName, ".")                            'ensure default ".zip" extension there if none provided
        If Not CBool(I) Then ZipName = ZipName & ".zip"
        OrgZipName = ZipName                                  'save as original, in case we later need to add more
        ZipPath = AddSlash(dst) & ZipName                     'set path to target zip file
        On Error Resume Next
        Set Ts = fso.OpenTextFile(ZipPath, ForWriting, True)  'try opening it as text
        If Not CBool(Err.Number) Then                         'errors (this is what we are checking for, we will destroy)
          Ts.Close
          fso.DeleteFile ZipPath, False                       'delete the existing file
        End If
        If CBool(Err.Number) Then                             'errors encountered?
          LogDataAdd "Error Creating: " & ZipName            'yes, report it
          Call ExitXcopy
          CenterMsgBoxOnForm frmXcopy, "Error creating ZIP file: " & ZipName, vbOKOnly Or vbExclamation, "ZIP Create Error"
          Exit Function
        End If
        On Error GoTo 0
      Else
        LogDataAdd "Error Creating: " & ZipName & ". Cannot write ZIP to blank file."
        Call ExitXcopy
        CenterMsgBoxOnForm frmXcopy, "Cannot write ZIP to blank file.", vbOKOnly Or vbExclamation, "Zip Error"
        Exit Function
      End If
    End If
'
' if we will be zipping, set index to strip base path from for relative path processing
'
    If CBool(Len(ZipPath)) Then
      If Right$(src, 1) = "\" Then
        ZipStrip = InStrRev(src, "\", Len(src) - 1) + 1
      Else
        ZipStrip = Len(src) + 2
      End If
'
' now initialize zip object
'
      With .m_Zip
        .ZipFile = ZipPath                      'set path to ZIP file to create/update
          If ZipStrip < 5 Then
            S = Left$(src, ZipStrip - 1)        'if index < 5, keep trailing '\' (drivespec)
          Else
            S = Left$(src, ZipStrip - 2)        'else pull back below '\'
          End If
        .BasePath = S                           'set base path to parent of project folder
        .Encrypt = False                        'no encryption
        .AddComment = False                     'no comment
        .IncludeSystemAndHiddenFiles = True     'allow hidden and system files
        .StoreFolderNames = True                'keep track of folder and subfolder paths
        .ClearFileSpecs                         'init for new list (max 1022 files provided for)
        ZipInc = 0
      End With
    End If
  End With
'
' now we are finally ready to start copying files
'
  S = src                                       'save copy of source (call to copyfolders will change it
  With frmXcopy
    .lblStatus.Caption = "Accumulating data for processing..."  'init for counting
    ItemCnt = 0                                 'init counters
    FolderCnt = 0
    FileCnt = 0
    .lblFolders.Caption = "Folders counted: 0"  'init reports
    .lblFiles.Caption = "Files counted: 0"
    Call CopyFolders(S, dst, True)              'count items
    .lblStatus.Caption = vbNullString
    MaxCnt = ItemCnt                            'save item count
    OfFolders = " of " & Format(FolderCnt, "#,##0")
    OfFiles = " of " & Format(FileCnt, "#,##0")
    ItemCnt = 0                                 're-init counteres
    FolderCnt = 0
    FileCnt = 0
    .lblFolders.Caption = "Folders copied: 0" & OfFolders   'prepare for copying
    .lblFiles.Caption = "Files copied: 0" & OfFiles
    With .lblPcent
      .Caption = "0%"                           'init 0% and show it
      .Visible = True
    End With
    If .WindowState = vbMinimized Then
      .Caption = "Xcpy [0%] from " & .txtFrom.Text
      .Refresh
    End If
    With .ProgressBar1
      .Value = 0
      .Visible = True                           'show process bar
      Call CopyFolders(src, dst, False)         'build/copy/move data
      .Visible = False                          'hide progress
    End With
    S = "XCPY " & GetAppVersion()
    If .Caption <> S Then .Caption = S
    With .tmrDelay
      .Enabled = True                           'enable-1-second timer
      Do While .Enabled                         'wait until it turns itself off
        DoEvents
      Loop
    End With
    
    Dta = "; " & .lblFolders & "; " & .lblFiles
    .lblPcent.Visible = False                   'hide % report
    .lblFolders.Caption = vbNullString          'clear other report fields
    .lblFiles.Caption = vbNullString
    .lblStatus.Caption = vbNullString
  End With
'
' if zipping, see if anything left to process
'
  If CBool(Len(ZipPath)) And Not Cancel Then
    S = Trim$(frmXcopy.txtTo)                   'get copy of destination
    With frmXcopy.m_Zip
      If CBool(.FileSpecCount) Then             'if files present to zip...
        frmXcopy.lblStatus.Caption = "Building " & ZipName & "..."
        .Zip                                    'zip the data and create the ZIP file
        frmXcopy.lblStatus.Caption = vbNullString
        If .Success Then                        'if there were no errors...
          CenterMsgBoxOnForm frmXcopy, "ZIP Backup created in """ & S & """", _
                             vbOKOnly Or vbInformation, "Zip File Created"
        Else                                    'if there were error...
          
          LogDataAdd "Error Creating: " & ZipName & " in """ & S & """"
          CenterMsgBoxOnForm frmXcopy, "There was an error creating " & ZipName & " in """ & S & """", _
                             vbOKOnly Or vbInformation, "Zip File Error"
        End If
      ElseIf .Success Then                        'if there were no errors...
        CenterMsgBoxOnForm frmXcopy, "ZIP Backup created in """ & S & """", _
                           vbOKOnly Or vbInformation, "Zip File(s) Created"
      End If
    End With
  End If
'
' write stuff to log file if present
'
  Dt = Now
  frmXcopy.lblEndTime.Caption = Format(Dt, "HH:MM:SS")
  ElapsedTime = Dt - StartTime               'update elapsed time
  S = Format(ElapsedTime, "HH:MM:SS")
  If CBool(StrComp(S, LastTime, vbTextCompare)) Then
    frmXcopy.lblElapsed.Caption = S
  End If
  
  S = CStr(Dt)
  LogDataAdd "End" & LogHead & S & Dta
  If CBool(Len(LogFile)) Then
    If CBool(LogData.Count) Then
      I = FreeFile()
      On Error GoTo 0
      If AppendLog Then
        Open LogFile For Append As #I
      Else
        Open LogFile For Output As #I
      End If
      If CBool(Err.Number) Then
        CenterMsgBoxOnForm frmXcopy, "Could not create log file: " & LogFile, vbOKOnly Or vbExclamation, "Log File Error"
      Else
        With LogData
          For Idx = 1 To .Count
            Print #I, .Item(Idx)
          Next Idx
          Close #I
        End With
      End If
    End If
  End If
'
' purge and remvoe the error list
'
  With LogData
    Xcopy = .Count - 1                                           'allow for Begin
    If Left$(.Item(.Count), 4) = "End " Then Xcopy = Xcopy - 1   'allow for End
    Do While .Count
      .Remove 1
    Loop
  End With
End Function

'*******************************************************************************
' Function Name     : IsDriveReady
' Purpose           : Make shure a drivespec is actually mounted
'*******************************************************************************
Public Function IsDriveReady(Path As String) As Boolean
  Dim I As Long
  Dim drv As Drive
  
  If Mid$(Path, 2, 1) = ":" Then     'std drivespec?
    Path = Left$(Path, 3)            'yes
  ElseIf Left$(Path, 2) = "\\" Then  'netpath?
    I = InStr(3, Path, "\")          'yes, so get network Source
    If CBool(I) Then
      I = InStr(I + 1, Path, "\")    'find netdrivespec
      If CBool(I) Then
        Path = Left$(Path, I)        'netdrivespec
      Else
        Path = vbNullString          'should never happen
      End If
    Else
      Path = vbNullString            'should never happen
    End If
  Else
    Path = vbNullString              'should never happen
  End If
  If CBool(Len(Path)) Then           'something to check
    On Error Resume Next
    Set drv = fso.GetDrive(Path)
    If Not CBool(Err.Number) Then
      IsDriveReady = drv.IsReady     'set true if drive is ready
    End If
  End If
End Function

'*******************************************************************************
' Function Name     : CopyFolders
' Purpose           : Copy folder information
'*******************************************************************************
Private Function CopyFolders(src As String, dst As String, Optional CountData As Boolean = False)
  Dim CurSrcDirList As Collection
  Dim CurDstDirList As Collection
  Dim CurFileList As Collection
  Dim S As String, D As String, Bugger As String, Tsrc As String
  Dim Idx As Long
  Dim I As Integer
  
  If Cancel Then Exit Function
  
  Set CurSrcDirList = New Collection              'set aside collections
  Set CurDstDirList = New Collection              'set aside collections
  Set CurFileList = New Collection
  
  If Right$(src, 1) = "\" Then
    src = RemoveSlash(src)
    I = InStrRev(src, "\")
    If CBool(I) Then
      CurSrcDirList.Add src                       'add directories we encounter
      CurDstDirList.Add RemoveSlash(dst) & Mid$(src, I)
      src = vbNullString                          'only copying src folder, so nothing else
    End If
  End If
  DoEvents
  
  If CBool(Len(src)) Then
    If InStr(1, src, "*") Or InStr(1, src, "?") Then
      I = InStrRev(src, "\")
      Bugger = Mid$(src, I)
      src = Left$(src, I - 1)
    Else
      Bugger = "\*.*"
    End If
    Tsrc = src & Bugger
    
    If Not CBool(Len(ZipPath)) Then     'if not to a zip file...
      If Not fso.FolderExists(dst) Then
        D = RemoveSlash(dst)
        On Error Resume Next
        CheckDirs D                     'make sure destination path exists (build up to it)
        MkDir D                         'then create target itself
        On Error GoTo 0
      End If
    End If
    
    If IncSubDir Then                               'if including subdirs
      S = Dir$(Tsrc, vbDirectory Or vbHidden Or vbSystem) 'grab all
    Else
      S = Dir$(Tsrc, vbHidden Or vbSystem)          'else grab just files
    End If
    Do While Len(S)
      If Left$(S, 1) <> "." Then                    'ignore . and ..
        S = "\" & S
        Idx = GetAttr(src & S)                      'grab attributes
        If Idx And vbDirectory Then                 'directory?
          CurSrcDirList.Add src & S                 'add directories we encounter
          CurDstDirList.Add dst & S
        Else
          If NewOnly Then
            If Not fso.FileExists(dst & S) Then     'if destination files does not exist...
              If ModOnly Then                       'is src to be also copied only if modified?
                If Idx And vbArchive Then           'yes, so check for modified flag
                  CurFileList.Add src & S           'add if file is new and has been modified
                End If
              Else
                CurFileList.Add src & S             'else add if new and modified does not matter
              End If
            End If
          ElseIf ModOnly Then
            If Idx And vbArchive Then               'check for modified only
              CurFileList.Add src & S               'add only files
            End If
          Else
            CurFileList.Add src & S                 'add files
          End If
        End If
      End If
      S = Dir$()                                    'grab next
    Loop                                            'do everything there
  End If
'
' process file list. Copy all files in collection
'
  With CurFileList
    Do While .Count                       'any files?
      If Pausing Then Call PauseProcess
      If Cancel Then Exit Do
      If CountData Then
        FileCnt = FileCnt + 1
        ItemCnt = ItemCnt + 1
        With frmXcopy.lblFiles
          .Caption = "Files counted: " & Format(FileCnt, "#,##0")
          .Refresh
        End With
  
        ElapsedTime = Now - StartTime               'update elapsed time
        D = Format(ElapsedTime, "HH:MM:SS")
        If CBool(StrComp(D, LastTime, vbTextCompare)) Then
          LastTime = D
          With frmXcopy.lblElapsed
            .Caption = D
            .Refresh
          End With
        End If
      Else
        CopyFile .Item(1), dst              'copy one
      End If
      .Remove 1                           'remove current
    Loop                                  'do all
  End With
  Set CurFileList = Nothing
'
' if any directories, then recurse through them
'
  With CurSrcDirList
    Do While .Count                       'any left?
      If Pausing Then Call PauseProcess
      If Cancel Then Exit Do
      CopyFolders .Item(1) & "\*.*", CurDstDirList(1), CountData 'process directory
      ItemCnt = ItemCnt + 1
      FolderCnt = FolderCnt + 1
      With frmXcopy
        With .lblFolders
          If CountData Then
            .Caption = "Folders counted: " & Format(FolderCnt, "#,##0")
          Else
            .Caption = "Folders copied: " & Format(FolderCnt, "#,##0") & OfFolders
          End If
          .Refresh
        End With
        
        ElapsedTime = Now - StartTime               'update elapsed time
        S = Format(ElapsedTime, "HH:MM:SS")
        If CBool(StrComp(S, LastTime, vbTextCompare)) Then
          With frmXcopy.lblElapsed
            .Caption = S
            .Refresh
          End With
        End If
        
        If Not CountData Then
          Pcnt = Round((ItemCnt / MaxCnt * 100 + 0.5), 0)
          If Pcnt > 100 Then Pcnt = 100
          With .ProgressBar1
            If .Value <> Pcnt Then
              .Value = Pcnt
              With frmXcopy
                With .lblPcent
                  .Caption = CStr(Pcnt) & "%"
                  .Refresh
                End With
                If .WindowState = vbMinimized Then
                  S = "Xcpy [" & .lblPcent.Caption & "] from " & .txtFrom.Text
                Else
                  S = "XCPY " & GetAppVersion()
                End If
                If .Caption <> S Then
                  .Caption = S
                  .Refresh
                End If
              End With
            End If
          End With
        End If
      End With
      .Remove 1                           'remove current
      CurDstDirList.Remove 1
    Loop                                  'check for more
  End With
  Set CurSrcDirList = Nothing
  Set CurDstDirList = Nothing
End Function

'*******************************************************************************
' Subroutine Name   : CopyFile
' Purpose           : Copy a file from Src to dst. Errors to Log file
'*******************************************************************************
Private Sub CopyFile(src As String, dst As String)
  Dim I As Integer
  Dim D As String, S As String, Fl As String
  
  DoEvents
  If Not CBool(Len(ZipPath)) Then
    If Not fso.FolderExists(dst) Then
      D = RemoveSlash(dst)
      On Error Resume Next
      CheckDirs RemoveSlash(D)               'make sure destination path exists
      MkDir D
      On Error GoTo 0
    End If
  End If
  
  On Error Resume Next
  I = InStrRev(src, "\")
  D = RemoveSlash(dst) & Mid$(src, I)        'add filename to destination
  
  With frmXcopy
    With .lblStatus
      .Caption = src
      .Refresh
    End With
    If .WindowState = vbMinimized Then
      S = "Xcpy [" & .lblPcent.Caption & "] from " & .txtFrom.Text
    Else
      S = "XCPY " & GetAppVersion()
    End If
    If .Caption <> S Then
      .Caption = S
      .Refresh
    End If
  End With
  
  DoEvents
  On Error GoTo 0
  If CBool(Len(ZipPath)) Then
    With frmXcopy.m_Zip
      .AddFileSpec Mid$(src, ZipStrip)
      If .FileSpecCount = 1022 Then
        frmXcopy.lblStatus.Caption = "Building " & ZipName & "..."
        .Zip                                    'zip current file
        frmXcopy.lblStatus.Caption = vbNullString
        If Not .Success Then
          LogDataAdd "Error Creating: " & ZipName & " in """ & Trim$(frmXcopy.txtTo) & """"
          CenterMsgBoxOnForm frmXcopy, "There was an error creating " & ZipName & " in """ & Trim$(frmXcopy.txtTo) & """", _
                             vbOKOnly Or vbInformation, "Zip File Error"
          Cancel = True
          Exit Sub
        End If
        ZipInc = ZipInc + 1                     'bump naming increment
        I = InStrRev(OrgZipName, ".")
        ZipName = Left$(OrgZipName, I - 1) & "_" & CStr(ZipInc) & Mid$(OrgZipName, I)
        ZipPath = AddSlash(Trim$(frmXcopy.txtTo.Text)) & ZipName
        If fso.FileExists(ZipPath) Then         'if zip already exists...
          On Error Resume Next
          Call fso.DeleteFile(ZipPath, True)    'remove it
          Err.Clear
        End If
        .ZipFile = ZipPath                      'set path to ZIP file to create/update
        If ZipStrip < 5 Then
          D = Left$(src, ZipStrip - 1)
        Else
          D = Left$(src, ZipStrip - 2)
        End If
        .BasePath = D                           'set base path to parent of project folder
        .Encrypt = False                        'no encryption
        .AddComment = False                     'no comment
        .IncludeSystemAndHiddenFiles = True     'allow hidden and system files
        .StoreFolderNames = True                'keep track of folder and subfolder paths
        .ClearFileSpecs                         'init for new list
      End If
    End With
  Else
    On Error Resume Next                        'trap errors
'    fso.CopyFile src, D, True                   'copy stuff
    CopyCancel = 0                              'reset copy cancel flag
    If NT Then                                  'copy stuff for NT/XP...
      Call CopyFileEx(src, D, AddressOf CopyFileCallback, 0&, CopyCancel, 0&)
    Else
      Call CopyFileA(src, D, 0)                 'copy stuff for Win95/98/Me
    End If
  End If
  
  FileCnt = FileCnt + 1                         'count file
  ItemCnt = ItemCnt + 1                         'count item
  With frmXcopy
    With .lblFiles
      .Caption = "Files copied: " & Format(FileCnt, "#,##0") & OfFiles
      .Refresh
    End With
    With .ProgressBar1
      Pcnt = Round((ItemCnt / MaxCnt * 100 + 0.5), 0)
      If Pcnt > 100 Then Pcnt = 100
      If .Value <> Pcnt Then
        .Value = Pcnt
        With frmXcopy.lblPcent
          .Caption = CStr(Pcnt) & "%"
          .Refresh
        End With
      End If
    End With
  End With
  
  If CBool(Err.Number) Then
    LogDataAdd "Error Copying: " & src & " TO " & D
  Else
    If frmXcopy.chkRstArchive.Value = vbChecked Then
      I = GetAttr(src) And &H6F   'get file attributes
      If I And vbArchive Then     'archive bit set?
        I = I - vbArchive         'yes, remove from attributes
        On Error Resume Next
        Call SetAttr(src, I) 'reset attributes
        If CBool(Err.Number) Then
          LogDataAdd "Archive Error: Could not reset archive flag on: " & src
        End If
        I = GetAttr(D) And &H6F       'get Dest file attributes
        If I And vbArchive = 0 Then   'archive bit set?
          I = I + vbArchive
          On Error Resume Next
          Call SetAttr(D, I)          'reset attributes
          If CBool(Err.Number) Then
            LogDataAdd "Archive Error: Could not reset archive flag on: " & src
          End If
        End If
      End If
    End If
  End If
  On Error GoTo 0
  
  ElapsedTime = Now - StartTime               'update elapsed time
  D = Format(ElapsedTime, "HH:MM:SS")
  If CBool(StrComp(D, LastTime, vbTextCompare)) Then
    LastTime = D
    With frmXcopy.lblElapsed
      .Caption = D
      .Refresh
    End With
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ExitXcopy
' Purpose           : Remove temp drive connections, if any
'*******************************************************************************
Private Sub ExitXcopy()
  If Len(TempSrc) Then DisconnectNetDrive TempSrc
  If Len(TempDst) Then DisconnectNetDrive TempDst
  
  With frmXcopy.Animation1
    If .Visible Then
      .Close               'close animation
      .AutoPlay = False
      .Visible = False
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : LogDataAdd
' Purpose           : 'Add data to the log file
'*******************************************************************************
Public Sub LogDataAdd(Msg As String)
  Dim Cnt As Long
  Dim S As String
  
  S = Msg
  If CBool(Err.Number) Then
    S = S & vbCrLf & vbCrLf & "Error Number: " & CStr(Err.Number) & vbCrLf & "Description: " & Err.Description
  End If
  With LogData
    .Add S                                                  'add message
    Cnt = .Count - 1                                        'allow for Begin
    If Left$(.Item(.Count), 4) = "End " Then Cnt = Cnt - 1  'allow for End
  End With
  
  With frmXcopy.lblErrors
    .Caption = CStr(Cnt)                                    'set error count
    .Refresh                                                'ensure displayed
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : PauseProcess
' Purpose           : wait for a continue or cancel
'*******************************************************************************
Private Sub PauseProcess()
  Do While Pausing
    If Cancel Then Pausing = False
    DoEvents
  Loop
End Sub

'*******************************************************************************
' Subroutine Name   : CopyFileCallback
' Purpose           : Callback for copy process progress indicator (used by NT/XP)
'*******************************************************************************
Private Function CopyFileCallback( _
                   ByVal TotalFileSize As Currency, _
                   ByVal TotalBytesTransferred As Currency, _
                   ByVal StreamSize As Currency, _
                   ByVal StreamBytesTransferred As Currency, _
                   ByVal dwStreamNumber As Long, _
                   ByVal dwCallbackReason As Long, _
                   ByVal hSourceFile As Long, _
                   ByVal hDestinationFile As Long, _
                   ByRef lpData As Long) As Long

  Dim S As String
  
  Select Case dwCallbackReason
    Case CALLBACK_STREAM_SWITCH         'switched streams or new file
      CopyDirection = True              'reset copy direction
      CopyCntr = 0                      'and counter
      CopyMask = 1
      CopyFileCallback = PROGRESS_CONTINUE
      
    Case CALLBACK_CHUNK_FINISHED        'a chunk of data was copied
      With frmXcopy.lblStatus
        S = .Caption                        'get current file being copied
        If Right$(S, 1) <> "." Then         'if no dot following...
          CopyDirection = True              'reset copy direction
          CopyCntr = 0                      'and counter
        End If
        
        If CopyDirection Then           'if adding dots...
          CopyMask = CopyMask + 1       'bump mask
          If CopyMask Mod 31 = 0 Then   '32?
            CopyMask = 0                'reset mask if so
            CopyCntr = CopyCntr + 1     'bump counter
            If CopyCntr = 10 Then       'if at top, then reverse direction
              CopyDirection = False
            End If
            S = S & "."                 'add a dot
            .Caption = S                'update screen
            DoEvents
          End If
        Else                            'removing dots...
          CopyMask = CopyMask + 1       'bump mask
          If CopyMask Mod 31 = 0 Then   '32?
            CopyMask = 0                'reset mask
            .Caption = Left$(S, Len(S) - 1) 'update status
            DoEvents
          End If
        End If
      End With
      CopyFileCallback = PROGRESS_CONTINUE
  End Select
End Function
'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


Attribute VB_Name = "modLoadResource"
Option Explicit
'~modLoadResource.bas;
'Load a resource file and write it to a file for use or play a WAV resource
'************************************************************************
' modLoadResource - The LoadResource() function obtains a CUSTOM resource
'                   object from a project's Resource object (*.RES), and
'                   saves it to a specified filename. If a path is not
'                   specified with the target filename, then the file will
'                   be created in the system's TEMP directory. The returned
'                   path will be the full path to the file. It will return
'                   a blank if the file was not created, usually due to
'                   the target resource not existing in the resource object.
'                   An Example of the use for such an object is an AVI or
'                   animated Gif, which we do not want to load to the our
'                   install location. Note that the file should be deleted
'                   after we are through using it. Hence, on AVI files playing
'                   in an Animation control, we must break the control's link
'                   to the file with the Close method befor we can kill it.
'                   The following example demonstrates this:
'
'                   The PlaySoundResource() function plays a WAV file stored
'                   in a resource file under the SOUND heading. Simply supply
'                   the ID number stored in that resource file for your WAV
'                   file. It will also load under the CUSTOM heading if it
'                   cannot find the SOUND resource. Note that this does not work
'                   while running in the IDE (obviously, because we do not have
'                   an EXE running from which resources can be grabbed).
'EXAMPLE:
'Private AviPath As String                         'hold AVI filepath
'
'Private Sub Form_Load()
'  Call cmdStopAVI_Click                           'make sure stuff is stopped
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'  Call cmdStopAVI_Click                           'make sure stuff is stopped
'End Sub
'
'Private Sub cmdTestAvi_Click()                    'PLAY avi file
'  AviPath = LoadResource(101, "CopyFile.avi")     'create file from resource 101
'  If Len(AviPath) Then                            'if file created
'    With Me.aniAVI                                'play avifile
'      .AutoPlay = True                            'turn on autoplay
'      .Open AviPath                               'load avi file and play it
'    End With
'  End If
'  Me.cmdTestAvi.Enabled = False                   'disable test button
'  Me.cmdStopAVI.Enabled = True                    'enable stop button
'End Sub
'
'Private Sub cmdStopAVI_Click()                    'STOP play of avi file
'  With Me.aniAVI
'    .AutoPlay = False                             'turn off avi if playing
'    .Close                                        'close so we can kill temp file
'  End With
'  If Len(AviPath) Then
'    Kill AviPath                                  'delete the file
'    AviPath = vbNullstring                        'stomp path to death
'  End If
'  Me.cmdTestAvi.Enabled = True                    'enable test button
'  Me.cmdStopAVI.Enabled = False                   'disable stop button
'End Sub
'************************************************************************

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const MAX_PATH = 260
Private Const SND_SYNC = &H0              'play synchronously (return when done)
Private Const SND_ASYNC = &H1             'play asynchronously (return immediately)
Private Const SND_LOOP = &H8              'loop the sound until next sndPlaySound
Private Const SND_NODEFAULT = &H2         'silence not default, if sound not found
Private Const SND_MEMORY = &H4            'lpszSoundName points to a memory file

Public Function LoadResource(ResourceIDNum As Integer, FilePath As String) As String
  Dim bArray() As Byte                                  'array to hold binary object
  Dim fname As String                                   'copy of FilePath
  Dim I As Integer                                      'temp variable
  Dim strPath As String                                 'path built
  
  On Error Resume Next                                  'catch errors
  bArray = LoadResData(ResourceIDNum, "CUSTOM")
  If Err.Number = 0 Then                                'loaded object?
    fname = Trim$(FilePath)                             'yes, grab path
    If InStr(1, fname, "\") = 0 Then                    'check for pathspec
      If InStr(1, fname, ":") = 0 Then                  'no drivepath
        strPath = String$(MAX_PATH + 1, 0)              'init the path target
        I = CInt(GetTempPath(MAX_PATH, strPath))        'get the temp dir
        If I Then                                       'found?
          I = InStr(1, strPath, vbNullChar) - 1         'yes, find terminator
          strPath = Left$(strPath, I) & fname           'build path in TEMP
        Else
          strPath = fname                               'else stuff supplied path
        End If
      End If
    End If
    If Len(strPath) = 0 Then strPath = fname            'assume User supplied whole thing
    I = FreeFile
    Open strPath For Binary Access Write As #I          'can we create the file?
    If Err.Number = 0 Then                              'yes
      Put #I, , bArray                                  'so write data to file
      Close #I                                          'close the file
      LoadResource = strPath                            'return the target path
    End If
  End If
End Function

Public Function PlaySoundResource(ByVal SndID As Long, Optional Wait As Boolean = False, Optional Continuous As Boolean = False) As Boolean
  Dim m_snd() As Byte
  Dim flags As Long
  
  If SndID Then                                       'if ID specified
    On Error Resume Next
    m_snd = LoadResData(SndID, "WAVE")                'load from WAVE resource
    If CBool(Err.Number) Then
      On Error Resume Next
      m_snd = LoadResData(SndID, "SOUND")             'load from SOUND resource
    End If
    If CBool(Err.Number) Then
      On Error Resume Next
      m_snd = LoadResData(SndID, "CUSTOM")            'load from custom resource
    End If
    If Err.Number = 0 Then                            'found it
      flags = SND_MEMORY Or SND_NODEFAULT             'inti flag
      If Not Wait Then
        flags = flags Or SND_ASYNC                    'do not wait until done
        If Continuous Then flags = flags Or SND_LOOP  'continuously play
      End If
      PlaySoundResource = Not CBool(PlaySoundData(m_snd(0), 0&, flags))
    End If
  Else      'stop playing any sound
    flags = SND_MEMORY Or SND_NODEFAULT Or SND_ASYNC  'init for stopping
    ReDim m_snd(0) As Byte                            'dead data
    PlaySoundResource = Not CBool(PlaySoundData(m_snd(0), 0&, flags))
  End If
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


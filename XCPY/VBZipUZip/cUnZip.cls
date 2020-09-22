VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUnzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'~cUnZip.cls;modUnzip.bas;vbuzip10.dll;
'Unzip data from a zip file
' ======================================================================================
' Name:     vbAccelerator cUnzip class
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     1 January 2000
'
' Requires: Info-ZIP's Unzip32.DLL v5.40, renamed to vbuzip10.dll
'           mUnzip.bas
'
' Copyright © 2000 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Part of the implementation of cUnzip.cls, a class which gives a
' simple interface to Info-ZIP's excellent, free unzipping library
' (Unzip32.DLL).
'
' Private WithEvents m_cUnzip As cUnzip 'declare
'
' Set m_cUnzip = New cUnzip             'Set up unzipping object
'
' Set m_cUnzip = Nothing                'release resources
'
' SUPPORTED EVENTS:
'   Cancel                              'user cancel
'   OverwritePrompt                     'prompt for target over-write
'   PasswordRequest                     'password required
'   Progress                            'trigger as each file is extracted
'
' PROPERITES
'   ExtractOnlyNewer                    'extract only newer files
'   SpaceToUnderScore                   'conver spaces to underscores
'   PromptToOverwrite                   'if over-write prompts are to be given
'   MessageLevel
'   TestZip                             'see if zip file is intact
'   UseFolderNames                      'if folders are to be used
'   OverwriteExisting                   'if files should be over-written
'   Convert CRToCRLF                    'if to convert line terminators to PC format
'   CaseSensitiveFileNames              'if filenames are case sensitive
'
'   ZipFile                             'Zip file path
'   Filecount                           'get number of files in zip (do .Directory)
'   FileDirectory                       'get folder from target to place file
'   FileSize                            'get a file's size in bytes
'   FilePackedSize                      'size of packed file
'   FileCompressionRation
'   FileDate                            'file modify date
'   FileCRC
'   FileCompressionMethod               'method used to compress file
'   FileEncrypted                       'if file is encrypted
'   FileSelected                        'only selected files will be extracted
'
'   Directory                           'build internal directory of files in ZIP
'   Unzip                               'unzip file in list tagged selected
'   ReverseSlashes                      'flip slashes in an file
' ======================================================================================

Public Enum EUZMsgLevel
   euzAllMessages = 0
   euzPartialMessages = 1
   euzNoMessages = 2
End Enum
Public Enum EUZOverWriteResponse
   euzDoNotOverwrite = 100
   euzOverwriteThisFile = 102
   euzOverwriteAllFiles = 103
   euzOverwriteNone = 104
End Enum

Private m_sZipFile As String
Private m_sUnzipFolder As String
Private Type tZipContents
   sName As String
   sFolder As String
   lSize As Long
   lPackedSize As Long
   lFactor As Long
   sMethod As String
   dDate As Date
   lCrc As Long
   fEncryped As Boolean
   fSelected As Boolean
End Type
Private m_tZipContents() As tZipContents
Private m_iCount As Long
Private m_tDCL As DCLIST

Public Event Cancel(ByVal sMsg As String, ByRef bCancel As Boolean)
Public Event OverwritePrompt(ByVal sFIle As String, ByRef eResponse As EUZOverWriteResponse)
Public Event PasswordRequest(ByRef sPassword As String, ByRef bCancel As Boolean)
Public Event Progress(ByVal lCount As Long, ByVal sMsg As String)

Public Property Get ExtractOnlyNewer() As Boolean
   ExtractOnlyNewer = (m_tDCL.ExtractOnlyNewer <> 0)      ' 1=extract only newer
End Property
Public Property Let ExtractOnlyNewer(ByVal bState As Boolean)
   m_tDCL.ExtractOnlyNewer = Abs(bState)      ' 1=extract only newer
End Property
Public Property Get SpaceToUnderScore() As Boolean
   SpaceToUnderScore = (m_tDCL.SpaceToUnderScore <> 0)    ' 1=convert space to underscore
End Property
Public Property Let SpaceToUnderScore(ByVal bState As Boolean)
   m_tDCL.SpaceToUnderScore = Abs(bState)
End Property
Public Property Get PromptToOverwrite() As Boolean
   PromptToOverwrite = (m_tDCL.PromptToOverwrite <> 0)
End Property
Public Property Let PromptToOverwrite(ByVal bState As Boolean)
   m_tDCL.PromptToOverwrite = Abs(bState)
End Property
Public Property Get MessageLevel() As EUZMsgLevel
   MessageLevel = m_tDCL.fQuiet
End Property
Public Property Let MessageLevel(ByVal eLevel As EUZMsgLevel)
   m_tDCL.fQuiet = eLevel
End Property
Public Property Get TestZip() As Boolean
   TestZip = (m_tDCL.ntflag <> 0)
End Property
Public Property Let TestZip(ByVal bState As Boolean)
   m_tDCL.ntflag = Abs(bState)
End Property
Public Property Get UseFolderNames() As Boolean
   UseFolderNames = (m_tDCL.ndflag <> 0)
End Property
Public Property Let UseFolderNames(ByVal bState As Boolean)
   m_tDCL.ndflag = Abs(bState)
End Property
Public Property Get OverwriteExisting() As Boolean
   OverwriteExisting = (m_tDCL.noflag <> 0)
End Property
Public Property Let OverwriteExisting(ByVal bState As Boolean)
   m_tDCL.noflag = Abs(bState)
End Property
Public Property Get ConvertCRToCRLF() As Boolean
   ConvertCRToCRLF = (m_tDCL.naflag <> 0)
End Property
Public Property Let ConvertCRToCRLF(ByVal bState As Boolean)
   m_tDCL.naflag = Abs(bState)
End Property
Public Property Get CaseSensitiveFileNames() As Boolean
   CaseSensitiveFileNames = m_tDCL.C_flag
End Property
Public Property Let CaseSensitiveFileNames(ByVal bState As Boolean)
   m_tDCL.C_flag = Abs(bState)
End Property

Friend Sub DirectoryListAddFile( _
      ByVal sFileName As String, _
      ByVal sFolder As String, _
      ByVal dDate As Date, _
      ByVal lSize As Long, _
      ByVal lCrc As Long, _
      ByVal fEncrypted As Boolean, _
      ByVal lFactor As Long, _
      ByVal sMethod As String _
   )
   If (sFileName <> vbNullChar) And Len(sFileName) > 0 Then
      m_iCount = m_iCount + 1
      ReDim Preserve m_tZipContents(1 To m_iCount) As tZipContents
      With m_tZipContents(m_iCount)
         .sName = sFileName
         .sFolder = sFolder
         .dDate = dDate
         .lSize = lSize
         .lCrc = lCrc
         .lFactor = lFactor
         .sMethod = sMethod
         .fEncryped = fEncrypted
         ' Default to selected:
         .fSelected = True
      End With
   End If
End Sub
Friend Sub OverwriteRequest( _
      ByVal sFIle As String, _
      ByRef eResponse As EUZOverWriteResponse _
   )
   RaiseEvent OverwritePrompt(sFIle, eResponse)
End Sub
Friend Sub ProgressReport( _
      ByVal sMsg As String _
   )
   RaiseEvent Progress(1, sMsg)
End Sub
Friend Sub Service( _
      ByVal sMsg As String, _
      ByRef bCancel As Boolean _
   )
   RaiseEvent Cancel(sMsg, bCancel)
End Sub
      
Friend Sub PasswordRequest( _
      ByRef sPassword As String, _
      ByRef bCancel As Boolean _
   )
   RaiseEvent PasswordRequest(sPassword, bCancel)
End Sub
Public Property Get FileCount() As Long
   FileCount = m_iCount
End Property
Public Property Get Filename(ByVal lIndex As Long) As String
   Filename = m_tZipContents(lIndex).sName
End Property
Public Property Get FileDirectory(ByVal lIndex As Long) As String
   FileDirectory = m_tZipContents(lIndex).sFolder
End Property
Public Property Get FileSize(ByVal lIndex As Long) As Long
   FileSize = m_tZipContents(lIndex).lSize
End Property
Public Property Get FilePackedSize(ByVal lIndex As Long) As Long
   FilePackedSize = m_tZipContents(lIndex).lSize * m_tZipContents(lIndex).lFactor \ 100
End Property
Public Property Get FileCompressionRatio(ByVal lIndex As Long) As Long
   FileCompressionRatio = m_tZipContents(lIndex).lFactor
End Property
Public Property Get FileDate(ByVal lIndex As Long) As Date
   FileDate = m_tZipContents(lIndex).dDate
End Property
Public Property Get FileCRC(ByVal lIndex As Long) As Long
   FileCRC = m_tZipContents(lIndex).lCrc
End Property
Public Property Get FileCompressionMethod(ByVal lIndex As Long) As String
   FileCompressionMethod = m_tZipContents(lIndex).sMethod
End Property
Public Property Get FileEncrypted(ByVal lIndex As Long) As Boolean
   FileEncrypted = m_tZipContents(lIndex).fEncryped
End Property
Public Property Get FileSelected(ByVal lIndex As Long) As Boolean
   FileSelected = m_tZipContents(lIndex).fSelected
End Property
Public Property Let FileSelected(ByVal lIndex As Long, ByVal bState As Boolean)
   m_tZipContents(lIndex).fSelected = bState
End Property
Public Function Directory() As Long
Dim S(0 To 0) As String
   m_tDCL.lpszZipFN = m_sZipFile
   m_tDCL.lpszExtractDir = vbNullChar
   m_tDCL.nvflag = 1
   modUnzip.VBUnzip Me, m_tDCL, 0, S(), 0, S()
End Function
Public Function Unzip() As Boolean
Dim sInc() As String
Dim iIncCount As Long
Dim S() As String
Dim I As Long
   If (m_sZipFile <> "") Then
      If (m_iCount > 0) Then
         For I = 1 To m_iCount
            If (m_tZipContents(I).fSelected) Then
               iIncCount = iIncCount + 1
               ReDim Preserve sInc(1 To iIncCount) As String
               sInc(iIncCount) = ReverseSlashes(m_tZipContents(I).sFolder, m_tZipContents(I).sName)
            End If
         Next I
         If (iIncCount = m_iCount) Then
            iIncCount = 0
            ReDim sInc(0 To 0)
         End If
      End If
      m_tDCL.lpszZipFN = m_sZipFile
      m_tDCL.nvflag = 0
      m_tDCL.lpszExtractDir = m_sUnzipFolder
      Unzip = (modUnzip.VBUnzip(Me, m_tDCL, iIncCount, sInc(), 0, S()) <> 0)
   End If
End Function
Private Function ReverseSlashes( _
      ByVal sFolder As String, _
      ByVal sFIle As String _
   ) As String
Dim sOut As String
Dim iPos As Long, iLastPos As Long

   If Len(sFolder) > 0 And sFolder <> vbNullChar Then
      sOut = sFolder & "/" & sFIle
      iLastPos = 1
      Do
         iPos = InStr(iLastPos, sOut, "\")
         If (iPos <> 0) Then
            Mid$(sOut, iPos, 1) = "/"
            iLastPos = iPos + 1
         End If
      Loop While iPos <> 0
      ReverseSlashes = sOut
   Else
      ReverseSlashes = sFIle
   End If
End Function
Public Property Let UnzipFolder(ByVal sFolder As String)
   m_sUnzipFolder = sFolder
End Property
Public Property Get UnzipFolder() As String
   UnzipFolder = m_sUnzipFolder
   m_tDCL.lpszExtractDir = m_sUnzipFolder
End Property
Public Property Get ZipFile() As String
   ZipFile = m_sZipFile
End Property
Public Sub GetVersion( _
      ByRef lMajor As Long, _
      ByRef lMinor As Long, _
      ByRef lRevision As Long, _
      ByRef dDate As Date, _
      ByRef sExtraInfo As String _
   )
Dim tVer As UZPVER
Dim iPos As Long

   ' Set Version space
   With tVer
      .structlen = Len(tVer)
      .betalevel = Space$(9) & vbNullChar
      .date = Space$(19) & vbNullChar
      .zlib = Space$(9) & vbNullChar
   End With
    
   ' Get version
   UzpVersion2 tVer
   iPos = InStr(tVer.betalevel, vbNullChar)
   If (iPos > 1) Then
      tVer.betalevel = Left$(tVer.betalevel, iPos - 1)
   End If
   sExtraInfo = tVer.betalevel
  
   ' Date..
'''   Debug.Print tVer.date

   lMajor = tVer.windll.major
   lMinor = tVer.windll.minor
   lRevision = tVer.windll.patchlevel
      
End Sub

Public Property Let ZipFile(ByVal sFIle As String)
   m_sZipFile = sFIle
   m_iCount = 0
   Erase m_tZipContents
End Property



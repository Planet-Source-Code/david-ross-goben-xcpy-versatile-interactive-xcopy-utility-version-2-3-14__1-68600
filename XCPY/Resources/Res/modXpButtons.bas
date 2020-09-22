Attribute VB_Name = "modXpButtons"
Option Explicit
'~modXpButtons.bas;
'Compress VB programs with XP-Style buttons
'**************************************************************************************
' modXpButtons - This module, when emplemented as described below, will cause the
'                compiled EXE to use Xp-Style buttons, fields, etc. if the system is
'                and XP-type system.
'
' Two subroutines are provided:
' FormInitialize()       :  Initialize an application to use XP-style controls if allowed
' ShapePicBkgToTabStrip():  Shapes a picturebox control container to conform to the
'                           body shape of a 5.0 tabstrip control. This is useful if you
'                           have a picturebox to contain the controls for a particular tab
'                           and want the picturebox to fit snugly within the tabstrip's
'                           borders. Simply shape the picturebox to approximate the size
'                           as close as possible, so that reshapping does not make the
'                           final work area look lop-sided.
'
'E-Z Emplementation Procedures:
'-----------------------------
' Step 1: In the startup form for your program, call the FormInitialize() subroutine from
'         with it's Form_Initialize() event. This will tell the EXE loader that it should
'         examine an associated MANIFEST file (which will be built automatically by the
'         subroutine if it does not exist.
'
' Step 2: Compile your EXE and run it. Note that XP style buttons WILL NOT be shown in
'         the Development Mode, unless you do as noted below in the second note.
'-----------------------------
' NOTES: Buttons with their Style set to Graphical will not be changed. This is because
'        internally, they are no longer button classes, but picture boxes that VB treats
'        like buttons. Cool cheat, if you ask me.
'
' *****  The first time you run this and it creates the MANIFEST file, this will not
'        seem to work. It is because the MANIFEST file must exist when the application
'        starts. Re-running the application will work fine. Hence, the auto-create
'        feature should be seen as an aid to the developer, where it will most likely
'        be created running in the development mode, where you cannot see its final
'        effects anyway, unless you have a VB6.EXE.MANIFEST file in the VB98 folder
'        (Just copy the app-created MANIFEST file to the VB98 folder and rename it to
'        VB6.EXE.MANIFEST).
'
' *****  For distribution, you should include the associated MANIFEST file with the
'        EXE file.
'**************************************************************************************

'API stuff
Public Const ICC_INTERNET_CLASSES = &H800 'use this if you want to use the manifiest file
Public Const ICC_USEREX_CLASSES = &H200   'use this if you use a RES manifest file

Public Type INITCOMMONCONTROLSEX_TYPE
  dwSize As Long
  dwICC As Long
End Type
' Incorporate OS stuff so that we will not need the modGetOsType.bas module
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128 'Maintenance string for PSS usage
End Type
Private Const VER_PLATFORM_WIN32_NT = 2

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEX_TYPE) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'*******************************************************************************
' Subroutine Name   : FormInitialize
' Purpose           : Call this routine from your Form_Inialize() subroutine
'*******************************************************************************
Public Sub FormInitialize()
  Dim DirPath As String, FilePath As String
  Dim FileNum As Integer
  Dim OSV As OSVERSIONINFO
  Dim ComCtls As INITCOMMONCONTROLSEX_TYPE  ' identifies the control to register
  
  OSV.dwOSVersionInfoSize = Len(OSV)                      'set size of info block
  If CBool(GetVersionEx(OSV)) Then                        'get OS info
    If OSV.dwPlatformId = VER_PLATFORM_WIN32_NT Then      'WinNT type?
      If OSV.dwMajorVersion >= 5 Then                     'XP series (or later)?
'''
''' UNCOMMENT THE FOLLOWING IF YOU PREFER TO USE THE APPNAME.EXE.MANIFEST FILE FORMAT
'''
'''        DirPath = App.Path                                'get path to app
'''        If Right$(DirPath, 1) <> "\" Then DirPath = DirPath & "\"  'ensure trailing \
'''        FilePath = DirPath & App.EXEName & ".exe.MANIFEST" 'build manifest path
'''        If Len(Dir$(FilePath)) = 0 Then                   'file exists?
'''          FileNum = FreeFile(0)                           'so, so we are going to build it
'''          On Error Resume Next
'''          Open FilePath For Output As #FileNum            'open for writing
'''          If CBool(Err.Number) Then Exit Sub              'error (on CD?), so ignore
'''          Print #FileNum, _
'''            "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
'''            "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
'''            "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""x86"" name=""VB6 Project"" type=""win32"" />" & vbCrLf & _
'''            "<dependency>" & vbCrLf & _
'''            "<dependentAssembly>" & vbCrLf & _
'''            "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""x86"" publicKeyToken=""6595b64144ccf1df"" language=""*"" />" & vbCrLf & _
'''            "</dependentAssembly>" & vbCrLf & _
'''            "</dependency>" & vbCrLf & _
'''            "</assembly>"                                 'write manifest data
'''          Close #FileNum                                  'close up
'''          On Error GoTo 0
'''        End If
'
' now tell the system to use XP stuff
'
        With ComCtls
          .dwSize = Len(ComCtls)                  'structure size
          .dwICC = ICC_USEREX_CLASSES             'Tell system "I WANT XP Buttons" from the RES Manifiest
'''          .dwICC = ICC_INTERNET_CLASSES           'Tell system "I WANT XP Buttons" from a Manifest file
        End With
        Call InitCommonControlsEx(ComCtls)        'process info
      End If
    End If
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : ShapePicBkgToTabStrip
' Purpose           : Shape a picturebox Background to a 5.0 Tabstrip. This is
'                   : useful when you are placing a picturebox control container
'                   : on a Tabstrip, and want to be sure that the picturebox will
'                   : fill the tabstrip body.
'*******************************************************************************
Public Sub ShapePicBkgToTabStrip(pBackground As PictureBox, TbStrip As Object)
  pBackground.Left = TbStrip.Left + 15       'right of left border
  pBackground.Width = TbStrip.Width - 60     'keep inside right border
  pBackground.Top = TbStrip.Top + 330        'below top border
  pBackground.Height = TbStrip.Height - 375  'above bottom border
End Sub

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


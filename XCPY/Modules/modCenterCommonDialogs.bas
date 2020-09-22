Attribute VB_Name = "modCenterCommonDialogs"
Option Explicit
'~modCenterCommonDialogs.bas;
'Show Standard Common Dialog boxes, centered to form or screen
'*****************************************************************************
' modCenterCommonDialogs - Show CommonDialog boxes without using the CommonDialog
'                          control (API only). By default, these dialogs center
'                          on the parent form. Setting the optional CenterOnForm
'                          flag to FALSE will center the dialogs on the screen.
'EXAMPLES:
'
'Private Sub cmdOpen_Click()
'  Dim sOpen As SelectedFile
'  Dim Count As Integer
'  Dim FileList As String
'
'  On Error Resume Next
'  FileDialog.sFilter = "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
'  ' See Standard CommonDialog Flags for all options
'  FileDialog.Flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
'  FileDialog.sDlgTitle = "Show Open"
'  FileDialog.sInitDir = App.Path & "\"
'  sOpen = ShowOpen(Me)
'  If sOpen.bCanceled Then Exit Sub
'
'  FileList = "Directory : " & sOpen.sLastDirectory & vbCr
'  For Count = 1 To sOpen.nFilesSelected
'    FileList = FileList & sOpen.sFiles(Count) & vbCr
'  Next Count
'  Call MsgBox(FileList, vbOKOnly + vbInformation, "Show Open Selected")
'End Sub
'
'Private Sub cmdSave_Click()
'  Dim sSave As SelectedFile
'  Dim Count As Integer
'  Dim FileList As String
'
'  FileDialog.sFilter = "Text Files (*.txt)" & Chr$(0) & "*.sky" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"
'  ' See Standard CommonDialog Flags for all options
'  FileDialog.Flags = OFN_HIDEREADONLY
'  FileDialog.sDlgTitle = "Show Save"
'  FileDialog.sInitDir = App.Path & "\"
'
'  On Error Resume Next
'  sSave = ShowSave(Me)
'  If CBool(Err.Number) Then Exit Sub
'  FileList = "Directory : " & sSave.sLastDirectory & vbCr
'  For Count = 1 To sSave.nFilesSelected
'    FileList = FileList & sSave.sFiles(Count) & vbCr
'  Next Count
'  Call MsgBox(FileList, vbOKOnly + vbInformation, "Show Save Selected")
'End Sub
'
'Private Sub cmdFont_Click()
'  Dim sFont As SelectedFont
'
'  FontDialog.iPointSize = 12 * 10
'  sFont = ShowFont(Me, "Times New Roman")
'End Sub
'
'Private Sub cmdPrint_Click()
'  Call ShowPrinter(Me, cdlgReturnDC)
'End Sub
'
'Private Sub cmdColor_Click()
'  Dim sColor As SelectedColor
'  sColor = ShowColor(Me)
'End Sub
'
'Private Sub cmdPrintSetup_Click()
'  Call ShowPrinter(Me, cdlgReturnDC Or cdlgPrintSetup)
'End Sub
'*****************************************************************************

'*****************************************************************************
' API enums, types, flags, functions
'*****************************************************************************
Public Enum CDLGFlags
  cdlgAllPages = 0
  cdlgCollate = &O10
  cdlgDisablePrintToFile = &H8000
  cdlgHelpButton = &H800
  cdlgHidePrintToFile = &H100000
  cdlgNoPageNums = &H8
  cdlgNoSelection = &H4
  cdlgNoWarning = &H80
  cdlgPageNums = &H2
  cdlgPrintSetup = &H40
  cdlgPrintToFile = &H20
  cdlgReturnDC = &H100
  cdlgReturnDefault = &H400
  cdlgReturnIC = &H200
  cdlgSelection = &H1
  CdlgUseDevModeCopies = &H40000
End Enum

Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Const GWL_HINSTANCE = (-6)
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10
Const HCBT_ACTIVATE = 5
Const WH_CBT = 5

Dim hHook As Long

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Declare Function GetShortPathname Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256

Public Const LF_FACESIZE = 32

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Public Type OPENFILENAME
  nStructSize As Long
  hwndOwner As Long
  hInstance As Long
  sFilter As String
  sCustomFilter As String
  nCustFilterSize As Long
  nFilterIndex As Long
  sFIle As String
  nFileSize As Long
  sFileTitle As String
  nTitleSize As Long
  sInitDir As String
  sDlgTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExt As Integer
  sDefFileExt As String
  nCustDataSize As Long
  fnHook As Long
  sTemplateName As String
End Type

Type NMHDR
  hwndFrom As Long
  idfrom As Long
  code As Long
End Type

Type OFNOTIFY
  hdr As NMHDR
  lpOFN As OPENFILENAME
  pszFile As String        '  May be NULL
End Type

Type CHOOSECOLORS
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type CHOOSEFONTS
  lStructSize As Long
  hwndOwner As Long          '  caller's window handle
  hDC As Long                '  printer DC/IC or NULL
  lpLogFont As Long          '  ptr. to a LOGFONT struct
  iPointSize As Long         '  10 * size in points of selected font
  flags As Long              '  enum. type flags
  rgbColors As Long          '  returned text color
  lCustData As Long          '  data passed to hook fn.
  lpfnHook As Long           '  ptr. to hook function
  lpTemplateName As String     '  custom template name
  hInstance As Long          '  instance handle of.EXE that
  lpszStyle As String          '  return the style field here
  nFontType As Integer          '  same value reported to the EnumFonts
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long           '  minimum pt size allowed &
  nSizeMax As Long           '  max pt size allowed if
End Type

Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100

Public Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT

Public Const CF_SCREENFONTS = &H1
Public Const CF_PRINTERFONTS = &H2
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_SHOWHELP = &H4&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_USESTYLE = &H80&
Public Const CF_EFFECTS = &H100&
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_WYSIWYG = &H8000 '  must also have CF_SCREENFONTS CF_PRINTERFONTS
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_TTONLY = &H40000
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOVERTFONTS = &H1000000

Public Const SIMULATED_FONTTYPE = &H8000
Public Const PRINTER_FONTTYPE = &H4000
Public Const SCREEN_FONTTYPE = &H2000
Public Const BOLD_FONTTYPE = &H100
Public Const ITALIC_FONTTYPE = &H200
Public Const REGULAR_FONTTYPE = &H400

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"

Public Const CD_LBSELNOITEMS = -1
Public Const CD_LBSELCHANGE = 0
Public Const CD_LBSELSUB = 1
Public Const CD_LBSELADD = 2

Type PRINTDLGS
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  hDC As Long
  flags As Long
  nFromPage As Integer
  nToPage As Integer
  nMinPage As Integer
  nMaxPage As Integer
  nCopies As Integer
  hInstance As Long
  lCustData As Long
  lpfnPrintHook As Long
  lpfnSetupHook As Long
  lpPrintTemplateName As String
  lpSetupTemplateName As String
  hPrintTemplate As Long
  hSetupTemplate As Long
End Type

Public Const PD_ALLPAGES = &H0
Public Const PD_SELECTION = &H1
Public Const PD_PAGENUMS = &H2
Public Const PD_NOSELECTION = &H4
Public Const PD_NOPAGENUMS = &H8
Public Const PD_COLLATE = &H10
Public Const PD_PRINTTOFILE = &H20
Public Const PD_PRINTSETUP = &H40
Public Const PD_NOWARNING = &H80
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNIC = &H200
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_SHOWHELP = &H800
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000

Type DEVNAMES
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
End Type

Public Const DN_DEFAULTPRN = &H1

Public Type SelectedFile
  nFilesSelected As Integer
  sFiles() As String
  sLastDirectory As String
  bCanceled As Boolean
End Type

Public Type SelectedColor
  oSelectedColor As OLE_COLOR
  bCanceled As Boolean
End Type

Public Type SelectedFont
  sSelectedFont As String
  bCanceled As Boolean
  bBold As Boolean
  bItalic As Boolean
  nSize As Integer
  bUnderline As Boolean
  bStrikeOut As Boolean
  lColor As Long
  sFaceName As String
End Type

Public FileDialog As OPENFILENAME
Public ColorDialog As CHOOSECOLORS
Public FontDialog As CHOOSEFONTS
Public PrintDialog As PRINTDLGS
Dim ParenthWnd As Long

'*******************************************************************************
' Function Name     : ShowOpen
' Purpose           : Show Open Dialog
'*******************************************************************************
Public Function ShowOpen(Frm As Form, Optional ByVal centerForm As Boolean = True) As SelectedFile
  Dim Ret As Long
  Dim Count As Integer
  Dim fileNameHolder As String
  Dim LastCharacter As Integer
  Dim NewCharacter As Integer
  Dim tempFiles(1 To 200) As String
  Dim hInst As Long
  Dim Thread As Long
  Dim FhWnd As Long
  
  FhWnd = Frm.hwnd
  ParenthWnd = FhWnd
  FileDialog.nStructSize = Len(FileDialog)
  FileDialog.hwndOwner = FhWnd
  FileDialog.sFileTitle = Space$(2048)
  FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
  FileDialog.sFIle = FileDialog.sFIle & Space$(2047) & Chr$(0)
  FileDialog.nFileSize = Len(FileDialog.sFIle)
  
  FileDialog.flags = OFS_FILE_OPEN_FLAGS
  
  'Set up the CBT hook
  hInst = GetWindowLong(FhWnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()
  If centerForm = True Then
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
  Else
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
  End If
  
  Ret = GetOpenFileName(FileDialog)

  If Ret Then
    If Trim$(FileDialog.sFileTitle) = "" Then
      LastCharacter = 0
      Count = 0
      While ShowOpen.nFilesSelected = 0
        NewCharacter = InStr(LastCharacter + 1, FileDialog.sFIle, Chr$(0), vbTextCompare)
        If Count > 0 Then
          tempFiles(Count) = Mid(FileDialog.sFIle, LastCharacter + 1, NewCharacter - LastCharacter - 1)
        Else
          ShowOpen.sLastDirectory = Mid(FileDialog.sFIle, LastCharacter + 1, NewCharacter - LastCharacter - 1)
        End If
        Count = Count + 1
        If InStr(NewCharacter + 1, FileDialog.sFIle, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, FileDialog.sFIle, Chr$(0) & Chr$(0), vbTextCompare) Then
          tempFiles(Count) = Mid(FileDialog.sFIle, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFIle, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
          ShowOpen.nFilesSelected = Count
        End If
        LastCharacter = NewCharacter
      Wend
      ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
      For Count = 1 To ShowOpen.nFilesSelected
        ShowOpen.sFiles(Count) = tempFiles(Count)
      Next
    Else
      ReDim ShowOpen.sFiles(1 To 1)
      ShowOpen.sLastDirectory = Left$(FileDialog.sFIle, FileDialog.nFileOffset)
      ShowOpen.nFilesSelected = 1
      ShowOpen.sFiles(1) = Mid(FileDialog.sFIle, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFIle, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
    End If
    ShowOpen.bCanceled = False
    Exit Function
  Else
    ShowOpen.sLastDirectory = ""
    ShowOpen.nFilesSelected = 0
    ShowOpen.bCanceled = True
    Erase ShowOpen.sFiles
    Exit Function
  End If
End Function

'*******************************************************************************
' Function Name     : ShowSave
' Purpose           : Show Save Dialog
'*******************************************************************************
Public Function ShowSave(Frm As Form, Optional ByVal centerForm As Boolean = True) As SelectedFile
  Dim Ret As Long
  Dim hInst As Long
  Dim Thread As Long
  Dim FhWnd As Long
  
  FhWnd = Frm.hwnd
  ParenthWnd = FhWnd
  FileDialog.nStructSize = Len(FileDialog)
  FileDialog.hwndOwner = FhWnd
  FileDialog.sFileTitle = Space$(2048)
  FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
  FileDialog.sFIle = Space$(2047) & Chr$(0)
  FileDialog.nFileSize = Len(FileDialog.sFIle)
  
  If FileDialog.flags = 0 Then FileDialog.flags = OFS_FILE_SAVE_FLAGS
  
  'Set up the CBT hook
  hInst = GetWindowLong(FhWnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()
  If centerForm = True Then
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
  Else
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
  End If
  
  Ret = GetSaveFileName(FileDialog)
  ReDim ShowSave.sFiles(1)

  If Ret Then
    ShowSave.sLastDirectory = Left$(FileDialog.sFIle, FileDialog.nFileOffset)
    ShowSave.nFilesSelected = 1
    ShowSave.sFiles(1) = Mid(FileDialog.sFIle, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFIle, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
    ShowSave.bCanceled = False
    Exit Function
  Else
    ShowSave.sLastDirectory = ""
    ShowSave.nFilesSelected = 0
    ShowSave.bCanceled = True
    Erase ShowSave.sFiles
    Exit Function
  End If
End Function

'*******************************************************************************
' Function Name     : ShowColor
' Purpose           : Show Color Dialog
'*******************************************************************************
Public Function ShowColor(Frm As Form, Optional ByVal centerForm As Boolean = True) As SelectedColor
  Dim customcolors() As Byte  ' dynamic (resizable) array
  Dim I As Integer
  Dim Ret As Long
  Dim hInst As Long
  Dim Thread As Long
  Dim FhWnd As Long
  
  FhWnd = Frm.hwnd
  ParenthWnd = FhWnd
  If ColorDialog.lpCustColors = "" Then
    ReDim customcolors(0 To 16 * 4 - 1) As Byte  'resize the array
    For I = LBound(customcolors) To UBound(customcolors)
      customcolors(I) = 254 ' sets all custom colors to white
    Next I
    ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
  End If
  
  ColorDialog.hwndOwner = FhWnd
  ColorDialog.lStructSize = Len(ColorDialog)
  ColorDialog.flags = COLOR_FLAGS
  
  'Set up the CBT hook
  hInst = GetWindowLong(FhWnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()
  If centerForm = True Then
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
  Else
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
  End If
  
  Ret = ChooseColor(ColorDialog)
  If Ret Then
    ShowColor.bCanceled = False
    ShowColor.oSelectedColor = ColorDialog.rgbResult
    Exit Function
  Else
    ShowColor.bCanceled = True
    ShowColor.oSelectedColor = &H0&
    Exit Function
  End If
End Function

'*******************************************************************************
' Function Name     : ShowFont
' Purpose           : Show Font Dialog
'*******************************************************************************
Public Function ShowFont(Frm As Form, ByVal startingFontName As String, Optional ByVal centerForm As Boolean = True) As SelectedFont
  Dim Ret As Long
  Dim lfLogFont As LOGFONT
  Dim hInst As Long
  Dim Thread As Long
  Dim I As Integer
  Dim FhWnd As Long
  
  FhWnd = Frm.hwnd
  ParenthWnd = FhWnd
  FontDialog.nSizeMax = 0
  FontDialog.nSizeMin = 0
  FontDialog.nFontType = Screen.FontCount
  FontDialog.hwndOwner = FhWnd
  FontDialog.hDC = 0
  FontDialog.lpfnHook = 0
  FontDialog.lCustData = 0
  FontDialog.lpLogFont = VarPtr(lfLogFont)
  If FontDialog.iPointSize = 0 Then
    FontDialog.iPointSize = 10 * 10
  End If
  FontDialog.lpTemplateName = Space$(2048)
  FontDialog.rgbColors = RGB(0, 255, 255)
  FontDialog.lStructSize = Len(FontDialog)
  
  If FontDialog.flags = 0 Then
    FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
  End If
  
  For I = 0 To Len(startingFontName) - 1
    lfLogFont.lfFaceName(I) = Asc(Mid(startingFontName, I + 1, 1))
  Next
  
  'Set up the CBT hook
  hInst = GetWindowLong(FhWnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()
  If centerForm = True Then
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
  Else
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
  End If
  
  Ret = ChooseFont(FontDialog)
      
  If Ret Then
    ShowFont.bCanceled = False
    ShowFont.bBold = IIf(lfLogFont.lfWeight > 400, 1, 0)
    ShowFont.bItalic = lfLogFont.lfItalic
    ShowFont.bStrikeOut = lfLogFont.lfStrikeOut
    ShowFont.bUnderline = lfLogFont.lfUnderline
    ShowFont.lColor = FontDialog.rgbColors
    ShowFont.nSize = FontDialog.iPointSize / 10
    For I = 0 To 31
      ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr(lfLogFont.lfFaceName(I))
    Next

    ShowFont.sSelectedFont = Mid(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, Chr(0)) - 1)
    Exit Function
  Else
    ShowFont.bCanceled = True
    Exit Function
  End If
End Function

'*******************************************************************************
' Function Name     : ShowPrinter
' Purpose           : Show Print Dialog
'*******************************************************************************
Public Function ShowPrinter(Frm As Form, flags As CDLGFlags, Optional ByVal centerForm As Boolean = True) As Long
  Dim hInst As Long
  Dim Thread As Long
  Dim FhWnd As Long
  
  FhWnd = Frm.hwnd
  ParenthWnd = FhWnd
  PrintDialog.hwndOwner = FhWnd
  PrintDialog.lStructSize = Len(PrintDialog)
  
  'Set up the CBT hook
  hInst = GetWindowLong(FhWnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()
  If centerForm = True Then
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
  Else
    hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
  End If
  PrintDialog.lStructSize = Len(PrintDialog)
  PrintDialog.flags = flags
  ShowPrinter = PrintDlg(PrintDialog)
End Function

'*******************************************************************************
' Function Name     : WinProcCenterScreen
' Purpose           : Center Dialog on Screen
'*******************************************************************************
Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim rectForm As RECT, rectMsg As RECT
  Dim x As Long, Y As Long
  
  If lMsg = HCBT_ACTIVATE Then
    'Show the MsgBox at a fixed location (0,0)
    GetWindowRect wParam, rectMsg
    x = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.Left) / 2
    Y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.Top) / 2
    Debug.Print "Screen " & Screen.Height / 2
    Debug.Print "MsgBox " & (rectMsg.Right - rectMsg.Left) / 2
    SetWindowPos wParam, 0, x, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    'Release the CBT hook
    UnhookWindowsHookEx hHook
  End If
  WinProcCenterScreen = False
End Function

'*******************************************************************************
' Function Name     : WinProcCenterForm
' Purpose           : Center dialog on Form
'*******************************************************************************
Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim rectForm As RECT, rectMsg As RECT
  Dim x As Long, Y As Long
  
  'On HCBT_ACTIVATE, show the MsgBox centered over Form1
  If lMsg = HCBT_ACTIVATE Then
    'Get the coordinates of the form and the message box so that
    'you can determine where the center of the form is located
    GetWindowRect ParenthWnd, rectForm
    GetWindowRect wParam, rectMsg
    x = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
    Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
    'Position the msgbox
    SetWindowPos wParam, 0, x, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    'Release the CBT hook
    UnhookWindowsHookEx hHook
   End If
   WinProcCenterForm = False
End Function
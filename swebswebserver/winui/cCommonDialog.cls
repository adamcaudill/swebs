VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CSEH: WinUI Custom
'***************************************************************************
'
' SWEBS/WinUI
'
' Copyright (c) 2003 Steve McMahon.
' Portions Copyright (c) 2003 Adam Caudill.
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'***************************************************************************

Option Explicit

Public Enum EErrorCommonDialog
    eeBaseCommonDialog = 13450  ' CommonDialog
End Enum

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" _
   Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
   Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
   Alias "GetFileTitleA" (ByVal szFile As String, _
   ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private Type TCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function ChooseColor Lib "COMDLG32.DLL" _
   Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public Enum EChooseColor
    CC_RGBInit = &H1
    CC_FullOpen = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
    ' Win95 only
    CC_SolidColor = &H80
    CC_AnyColor = &H100
    ' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum
Private Declare Function GetSysColor Lib "USER32" (ByVal nIndex As Long) As Long

Private Type TCHOOSEFONT
    lStructSize As Long         ' Filled with UDT size
    hWndOwner As Long           ' Caller's window handle
    hdc As Long                 ' Printer DC/IC or NULL
    lpLogFont As Long           ' Pointer to LOGFONT
    iPointSize As Long          ' 10 * size in points of font
    flags As Long               ' Type flags
    rgbColors As Long           ' Returned text color
    lCustData As Long           ' Data passed to hook function
    lpfnHook As Long            ' Pointer to hook function
    lpTemplateName As Long      ' Custom template name
    hInstance As Long           ' Instance handle for template
    lpszStyle As String         ' Return style field
    nFontType As Integer        ' Font type bits
    iAlign As Integer           ' Filler
    nSizeMin As Long            ' Minimum point size allowed
    nSizeMax As Long            ' Maximum point size allowed
End Type
Private Declare Function ChooseFont Lib "COMDLG32" _
   Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long

Private Const LF_FACESIZE = 32
Private Type LOGFONT
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

Public Enum EChooseFont
    CF_ScreenFonts = &H1
    CF_PrinterFonts = &H2
    CF_BOTH = &H3
    CF_FontShowHelp = &H4
    CF_UseStyle = &H80
    CF_EFFECTS = &H100
    CF_AnsiOnly = &H400
    CF_NoVectorFonts = &H800
    CF_NoOemFonts = CF_NoVectorFonts
    CF_NoSimulations = &H1000
    CF_LimitSize = &H2000
    CF_FixedPitchOnly = &H4000
    CF_WYSIWYG = &H8000&  ' Must also have ScreenFonts And PrinterFonts
    CF_ForceFontExist = &H10000
    CF_ScalableOnly = &H20000
    CF_TTOnly = &H40000
    CF_NoFaceSel = &H80000
    CF_NoStyleSel = &H100000
    CF_NoSizeSel = &H200000
    ' Win95 only
    CF_SelectScript = &H400000
    CF_NoScriptSel = &H800000
    CF_NoVertFonts = &H1000000

    CF_InitToLogFontStruct = &H40
    CF_Apply = &H200
    CF_EnableHook = &H8
    CF_EnableTemplate = &H10
    CF_EnableTemplateHandle = &H20
    CF_FontNotSupported = &H238
End Enum

' These are extra nFontType bits that are added to what is returned to the
' EnumFonts callback routine

Public Enum EFontType
    Simulated_FontType = &H8000&
    Printer_FontType = &H4000
    Screen_FontType = &H2000
    Bold_FontType = &H100
    Italic_FontType = &H200
    Regular_FontType = &H400
End Enum

Private Type TPRINTDLG
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
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
    lpPrintTemplateName As Long
    lpSetupTemplateName As Long
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'  DEVMODE collation selections
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

Private Declare Function PrintDlg Lib "COMDLG32.DLL" _
   Alias "PrintDlgA" (prtdlg As TPRINTDLG) As Integer

Public Enum EPrintDialog
    PD_ALLPAGES = &H0
    PD_SELECTION = &H1
    PD_PAGENUMS = &H2
    PD_NOSELECTION = &H4
    PD_NOPAGENUMS = &H8
    PD_COLLATE = &H10
    PD_PRINTTOFILE = &H20
    PD_PRINTSETUP = &H40
    PD_NOWARNING = &H80
    PD_RETURNDC = &H100
    PD_RETURNIC = &H200
    PD_RETURNDEFAULT = &H400
    PD_SHOWHELP = &H800
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLESETUPTEMPLATE = &H8000&
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    PD_DISABLEPRINTTOFILE = &H80000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
End Enum

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DevMode
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

' New Win95 Page Setup dialogs are up to you
Private Type POINTL
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    TOp As Long
    Right As Long
    Bottom As Long
End Type

Private Type TPAGESETUPDLG
    lStructSize                 As Long
    hWndOwner                   As Long
    hDevMode                    As Long
    hDevNames                   As Long
    flags                       As Long
    ptPaperSize                 As POINTL
    rtMinMargin                 As RECT
    rtMargin                    As RECT
    hInstance                   As Long
    lCustData                   As Long
    lpfnPageSetupHook           As Long
    lpfnPagePaintHook           As Long
    lpPageSetupTemplateName     As Long
    hPageSetupTemplate          As Long
End Type

' EPaperSize constants same as vbPRPS constants
Public Enum EPaperSize
    epsLetter = 1          ' Letter, 8 1/2 x 11 in.
    epsLetterSmall         ' Letter Small, 8 1/2 x 11 in.
    epsTabloid             ' Tabloid, 11 x 17 in.
    epsLedger              ' Ledger, 17 x 11 in.
    epsLegal               ' Legal, 8 1/2 x 14 in.
    epsStatement           ' Statement, 5 1/2 x 8 1/2 in.
    epsExecutive           ' Executive, 7 1/2 x 10 1/2 in.
    epsA3                  ' A3, 297 x 420 mm
    epsA4                  ' A4, 210 x 297 mm
    epsA4Small             ' A4 Small, 210 x 297 mm
    epsA5                  ' A5, 148 x 210 mm
    epsB4                  ' B4, 250 x 354 mm
    epsB5                  ' B5, 182 x 257 mm
    epsFolio               ' Folio, 8 1/2 x 13 in.
    epsQuarto              ' Quarto, 215 x 275 mm
    eps10x14               ' 10 x 14 in.
    eps11x17               ' 11 x 17 in.
    epsNote                ' Note, 8 1/2 x 11 in.
    epsEnv9                ' Envelope #9, 3 7/8 x 8 7/8 in.
    epsEnv10               ' Envelope #10, 4 1/8 x 9 1/2 in.
    epsEnv11               ' Envelope #11, 4 1/2 x 10 3/8 in.
    epsEnv12               ' Envelope #12, 4 1/2 x 11 in.
    epsEnv14               ' Envelope #14, 5 x 11 1/2 in.
    epsCSheet              ' C size sheet
    epsDSheet              ' D size sheet
    epsESheet              ' E size sheet
    epsEnvDL               ' Envelope DL, 110 x 220 mm
    epsEnvC3               ' Envelope C3, 324 x 458 mm
    epsEnvC4               ' Envelope C4, 229 x 324 mm
    epsEnvC5               ' Envelope C5, 162 x 229 mm
    epsEnvC6               ' Envelope C6, 114 x 162 mm
    epsEnvC65              ' Envelope C65, 114 x 229 mm
    epsEnvB4               ' Envelope B4, 250 x 353 mm
    epsEnvB5               ' Envelope B5, 176 x 250 mm
    epsEnvB6               ' Envelope B6, 176 x 125 mm
    epsEnvItaly            ' Envelope, 110 x 230 mm
    epsenvmonarch          ' Envelope Monarch, 3 7/8 x 7 1/2 in.
    epsEnvPersonal         ' Envelope, 3 5/8 x 6 1/2 in.
    epsFanfoldUS           ' U.S. Standard Fanfold, 14 7/8 x 11 in.
    epsFanfoldStdGerman    ' German Standard Fanfold, 8 1/2 x 12 in.
    epsFanfoldLglGerman    ' German Legal Fanfold, 8 1/2 x 13 in.
    epsUser = 256          ' User-defined
End Enum

' EPrintQuality constants same as vbPRPQ constants
Public Enum EPrintQuality
    epqDraft = -1
    epqLow = -2
    epqMedium = -3
    epqHigh = -4
End Enum

Public Enum EOrientation
    eoPortrait = 1
    eoLandscape
End Enum

Private Declare Function PageSetupDlg Lib "COMDLG32" _
   Alias "PageSetupDlgA" (lppage As TPAGESETUPDLG) As Boolean

Public Enum EPageSetup
    PSD_Defaultminmargins = &H0 ' Default (printer's)
    PSD_InWinIniIntlMeasure = &H0
    PSD_MINMARGINS = &H1
    PSD_MARGINS = &H2
    PSD_INTHOUSANDTHSOFINCHES = &H4
    PSD_INHUNDREDTHSOFMILLIMETERS = &H8
    PSD_DISABLEMARGINS = &H10
    PSD_DISABLEPRINTER = &H20
    PSD_NoWarning = &H80
    PSD_DISABLEORIENTATION = &H100
    PSD_ReturnDefault = &H400
    PSD_DISABLEPAPER = &H200
    PSD_ShowHelp = &H800
    PSD_EnablePageSetupHook = &H2000
    PSD_EnablePageSetupTemplate = &H8000&
    PSD_EnablePageSetupTemplateHandle = &H20000
    PSD_EnablePagePaintHook = &H40000
    PSD_DisablePagePainting = &H80000
End Enum

Public Enum EPageSetupUnits
    epsuInches
    epsuMillimeters
End Enum

' Common dialog errors

Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF

    CDERR_GENERALCODES = &H0
    CDERR_STRUCTSIZE = &H1
    CDERR_INITIALIZATION = &H2
    CDERR_NOTEMPLATE = &H3
    CDERR_NOHINSTANCE = &H4
    CDERR_LOADSTRFAILURE = &H5
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHOOK = &HB
    CDERR_REGISTERMSGFAIL = &HC

    PDERR_PRINTERCODES = &H1000
    PDERR_SETUPFAILURE = &H1001
    PDERR_PARSEFAILURE = &H1002
    PDERR_RETDEFFAILURE = &H1003
    PDERR_LOADDRVFAILURE = &H1004
    PDERR_GETDEVMODEFAIL = &H1005
    PDERR_INITFAILURE = &H1006
    PDERR_NODEVICES = &H1007
    PDERR_NODEFAULTPRN = &H1008
    PDERR_DNDMMISMATCH = &H1009
    PDERR_CREATEICFAILURE = &H100A
    PDERR_PRINTERNOTFOUND = &H100B
    PDERR_DEFAULTDIFFERENT = &H100C

    CFERR_CHOOSEFONTCODES = &H2000
    CFERR_NOFONTS = &H2001
    CFERR_MAXLESSTHANMIN = &H2002

    FNERR_FILENAMECODES = &H3000
    FNERR_SUBCLASSFAILURE = &H3001
    FNERR_INVALIDFILENAME = &H3002
    FNERR_BUFFERTOOSMALL = &H3003

    CCERR_CHOOSECOLORCODES = &H5000
End Enum

' Array of custom colors lasts for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean

Public Enum EPrintRange
    eprAll
    eprPageNumbers
    eprSelection
End Enum
Private m_lApiReturn As Long
Private m_lExtendedError As Long
Private m_dvmode As DevMode

Public Property Get APIReturn() As Long
        '<EhHeader>
        On Error GoTo APIReturn_Err
        '</EhHeader>

        'return object's APIReturn property
100     APIReturn = m_lApiReturn

        '<EhFooter>
        Exit Property

APIReturn_Err:
104     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.APIReturn", Erl, False
108     Resume Next
        '</EhFooter>
End Property

Public Property Get ExtendedError() As Long
        '<EhHeader>
        On Error GoTo ExtendedError_Err
        '</EhHeader>

        'return object's ExtendedError property
100     ExtendedError = m_lExtendedError

        '<EhFooter>
        Exit Property

ExtendedError_Err:
104     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.ExtendedError", Erl, False
108     Resume Next
        '</EhFooter>
End Property

#If fComponent Then
Private Sub Class_Initialize()
    InitColors
End Sub
#End If

Function VBGetOpenFileName(Filename As String, _
   Optional FileTitle As String, _
   Optional FileMustExist As Boolean = True, _
   Optional MultiSelect As Boolean = False, _
   Optional ReadOnly As Boolean = False, _
   Optional HideReadOnly As Boolean = False, _
   Optional Filter As String = "All (*.*)| *.*", _
   Optional FilterIndex As Long = 1, _
   Optional InitDir As String, _
   Optional DlgTitle As String, _
   Optional DefaultExt As String, _
   Optional Owner As Long = -1, _
   Optional flags As Long = 0) As Boolean
        '<EhHeader>
        On Error GoTo VBGetOpenFileName_Err
        '</EhHeader>

        Dim opfile As OPENFILENAME, s As String, afFlags As Long
    
100     m_lApiReturn = 0
104     m_lExtendedError = 0

108     With opfile

112         .lStructSize = Len(opfile)
    
            ' Add in specific flags and strip out non-VB flags
    
116         .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
               (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
               (-ReadOnly * OFN_READONLY) Or _
               (-HideReadOnly * OFN_HIDEREADONLY) Or _
               (flags And CLng(Not (OFN_ENABLEHOOK Or _
               OFN_ENABLETEMPLATE)))
            ' Owner can take handle of owning window

120         If Owner <> -1 Then .hWndOwner = Owner

            ' InitDir can take initial directory string
124         .lpstrInitialDir = InitDir
            ' DefaultExt can take default extension
128         .lpstrDefExt = DefaultExt
            ' DlgTitle can take dialog box title
132         .lpstrTitle = DlgTitle
    
            ' To make Windows-style filter, replace | and : with nulls
            Dim ch As String, i As Integer

136         For i = 1 To Len(Filter)

140             ch = Mid$(Filter, i, 1)

144             If ch = "|" Or ch = ":" Then

148                 s = s & vbNullChar

                Else

152                 s = s & ch

                End If

            Next

            ' Put double null at end
156         s = s & vbNullChar & vbNullChar
160         .lpstrFilter = s
164         .nFilterIndex = FilterIndex

            ' Pad file and file title buffers to maximum path
168         s = Filename & String$(MAX_PATH - Len(Filename), 0)
172         .lpstrFile = s
176         .nMaxFile = MAX_PATH
180         s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
184         .lpstrFileTitle = s
188         .nMaxFileTitle = MAX_FILE
            ' All other fields set to zero
    
192         m_lApiReturn = GetOpenFileName(opfile)

196         Select Case m_lApiReturn

                Case 1
                    ' Success
200                 VBGetOpenFileName = True
204                 Filename = StrZToStr(.lpstrFile)
208                 FileTitle = StrZToStr(.lpstrFileTitle)
212                 flags = .flags
                    ' Return the filter index
216                 FilterIndex = .nFilterIndex
                    ' Look up the filter the user selected and return that
220                 Filter = FilterLookup(.lpstrFilter, FilterIndex)

224                 If (.flags And OFN_READONLY) Then ReadOnly = True

228             Case 0
                    ' Cancelled
232                 VBGetOpenFileName = False
236                 Filename = ""
240                 FileTitle = ""
244                 flags = 0
248                 FilterIndex = -1
252                 Filter = ""

256             Case Else
                    ' Extended error
260                 m_lExtendedError = CommDlgExtendedError()
264                 VBGetOpenFileName = False
268                 Filename = ""
272                 FileTitle = ""
276                 flags = 0
280                 FilterIndex = -1
284                 Filter = ""

            End Select

        End With

        '<EhFooter>
        Exit Function

VBGetOpenFileName_Err:
288     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.VBGetOpenFileName", Erl, False
292     Resume Next
        '</EhFooter>
End Function

Private Function StrZToStr(s As String) As String
        '<EhHeader>
        On Error GoTo StrZToStr_Err
        '</EhHeader>

100     StrZToStr = Left$(s, lstrlen(s))

        '<EhFooter>
        Exit Function

StrZToStr_Err:
104     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.StrZToStr", Erl, False
108     Resume Next
        '</EhFooter>
End Function

Function VBGetSaveFileName(Filename As String, _
   Optional FileTitle As String, _
   Optional OverWritePrompt As Boolean = True, _
   Optional Filter As String = "All (*.*)| *.*", _
   Optional FilterIndex As Long = 1, _
   Optional InitDir As String, _
   Optional DlgTitle As String, _
   Optional DefaultExt As String, _
   Optional Owner As Long = -1, _
   Optional flags As Long) As Boolean
        '<EhHeader>
        On Error GoTo VBGetSaveFileName_Err
        '</EhHeader>
            
        Dim opfile As OPENFILENAME, s As String

100     m_lApiReturn = 0
104     m_lExtendedError = 0

108     With opfile

112         .lStructSize = Len(opfile)
    
            ' Add in specific flags and strip out non-VB flags
116         .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
               OFN_HIDEREADONLY Or _
               (flags And CLng(Not (OFN_ENABLEHOOK Or _
               OFN_ENABLETEMPLATE)))
            ' Owner can take handle of owning window

120         If Owner <> -1 Then .hWndOwner = Owner

            ' InitDir can take initial directory string
124         .lpstrInitialDir = InitDir
            ' DefaultExt can take default extension
128         .lpstrDefExt = DefaultExt
            ' DlgTitle can take dialog box title
132         .lpstrTitle = DlgTitle
    
            ' Make new filter with bars (|) replacing nulls and double null at end
            Dim ch As String, i As Integer

136         For i = 1 To Len(Filter)

140             ch = Mid$(Filter, i, 1)

144             If ch = "|" Or ch = ":" Then

148                 s = s & vbNullChar

                Else

152                 s = s & ch

                End If

            Next

            ' Put double null at end
156         s = s & vbNullChar & vbNullChar
160         .lpstrFilter = s
164         .nFilterIndex = FilterIndex

            ' Pad file and file title buffers to maximum path
168         s = Filename & String$(MAX_PATH - Len(Filename), 0)
172         .lpstrFile = s
176         .nMaxFile = MAX_PATH
180         s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
184         .lpstrFileTitle = s
188         .nMaxFileTitle = MAX_FILE
            ' All other fields zero
    
192         m_lApiReturn = GetSaveFileName(opfile)

196         Select Case m_lApiReturn

                Case 1
200                 VBGetSaveFileName = True
204                 Filename = StrZToStr(.lpstrFile)
208                 FileTitle = StrZToStr(.lpstrFileTitle)
212                 flags = .flags
                    ' Return the filter index
216                 FilterIndex = .nFilterIndex
                    ' Look up the filter the user selected and return that
220                 Filter = FilterLookup(.lpstrFilter, FilterIndex)

224             Case 0
                    ' Cancelled:
228                 VBGetSaveFileName = False
232                 Filename = ""
236                 FileTitle = ""
240                 flags = 0
244                 FilterIndex = 0
248                 Filter = ""

252             Case Else
                    ' Extended error:
256                 VBGetSaveFileName = False
260                 m_lExtendedError = CommDlgExtendedError()
264                 Filename = ""
268                 FileTitle = ""
272                 flags = 0
276                 FilterIndex = 0
280                 Filter = ""

            End Select

        End With

        '<EhFooter>
        Exit Function

VBGetSaveFileName_Err:
284     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.VBGetSaveFileName", Erl, False
288     Resume Next
        '</EhFooter>
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
        '<EhHeader>
        On Error GoTo FilterLookup_Err
        '</EhHeader>

        Dim iStart As Long, iEnd As Long, s As String
100     iStart = 1

104     If sFilters = "" Then Exit Function

        Do

            ' Cut out both parts marked by null character
108         iEnd = InStr(iStart, sFilters, vbNullChar)

112         If iEnd = 0 Then Exit Function

116         iEnd = InStr(iEnd + 1, sFilters, vbNullChar)

120         If iEnd Then

124             s = Mid$(sFilters, iStart, iEnd - iStart)

            Else

128             s = Mid$(sFilters, iStart)

            End If

132         iStart = iEnd + 1

136         If iCur = 1 Then

140             FilterLookup = s
                Exit Function

            End If

144         iCur = iCur - 1

148     Loop While iCur

        '<EhFooter>
        Exit Function

FilterLookup_Err:
152     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.FilterLookup", Erl, False
156     Resume Next
        '</EhFooter>
End Function

Function VBGetFileTitle(sFile As String) As String
        '<EhHeader>
        On Error GoTo VBGetFileTitle_Err
        '</EhHeader>

        Dim sFileTitle As String, cFileTitle As Integer

100     cFileTitle = MAX_PATH
104     sFileTitle = String$(MAX_PATH, 0)
108     cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)

112     If cFileTitle Then

116         VBGetFileTitle = ""

        Else

120         VBGetFileTitle = Left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)

        End If

        '<EhFooter>
        Exit Function

VBGetFileTitle_Err:
124     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.VBGetFileTitle", Erl, False
128     Resume Next
        '</EhFooter>
End Function

' ChooseColor wrapper
Function VBChooseColor(Color As Long, _
   Optional AnyColor As Boolean = True, _
   Optional FullOpen As Boolean = False, _
   Optional DisableFullOpen As Boolean = False, _
   Optional Owner As Long = -1, _
   Optional flags As Long) As Boolean
        '<EhHeader>
        On Error GoTo VBChooseColor_Err
        '</EhHeader>

        Dim chclr As TCHOOSECOLOR
100     chclr.lStructSize = Len(chclr)
    
        ' Color must get reference variable to receive result
        ' Flags can get reference variable or constant with bit flags
        ' Owner can take handle of owning window

104     If Owner <> -1 Then chclr.hWndOwner = Owner

        ' Assign color (default uninitialized value of zero is good default)
108     chclr.rgbResult = Color

        ' Mask out unwanted bits
        Dim afMask As Long
112     afMask = CLng(Not (CC_ENABLEHOOK Or _
           CC_ENABLETEMPLATE))
        ' Pass in flags
116     chclr.flags = afMask And (CC_RGBInit Or _
           IIf(AnyColor, CC_AnyColor, CC_SolidColor) Or _
           (-FullOpen * CC_FullOpen) Or _
           (-DisableFullOpen * CC_PreventFullOpen))

        ' If first time, initialize to white

120     If fNotFirst = False Then InitColors

124     chclr.lpCustColors = VarPtr(alCustom(0))
        ' All other fields zero
    
128     m_lApiReturn = ChooseColor(chclr)

132     Select Case m_lApiReturn

            Case 1
                ' Success
136             VBChooseColor = True
140             Color = chclr.rgbResult

144         Case 0
                ' Cancelled
148             VBChooseColor = False
152             Color = -1

156         Case Else
                ' Extended error
160             m_lExtendedError = CommDlgExtendedError()
164             VBChooseColor = False
168             Color = -1

        End Select

        '<EhFooter>
        Exit Function

VBChooseColor_Err:
172     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.VBChooseColor", Erl, False
176     Resume Next
        '</EhFooter>
End Function

Private Sub InitColors()
        '<EhHeader>
        On Error GoTo InitColors_Err
        '</EhHeader>

        Dim i As Integer
        ' Initialize with first 16 system interface colors

100     For i = 0 To 15

104         alCustom(i) = GetSysColor(i)

        Next

108     fNotFirst = True

        '<EhFooter>
        Exit Sub

InitColors_Err:
112     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.InitColors", Erl, False
116     Resume Next
        '</EhFooter>
End Sub

' Property to read or modify custom colors (use to save colors in registry)
Public Property Get CustomColor(i As Integer) As Long
        '<EhHeader>
        On Error GoTo CustomColor_Err
        '</EhHeader>

        ' If first time, initialize to white

100     If fNotFirst = False Then InitColors

104     If i >= 0 And i <= 15 Then

108         CustomColor = alCustom(i)

        Else

112         CustomColor = -1

        End If

        '<EhFooter>
        Exit Property

CustomColor_Err:
116     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.CustomColor", Erl, False
120     Resume Next
        '</EhFooter>
End Property

Public Property Let CustomColor(i As Integer, iValue As Long)
        '<EhHeader>
        On Error GoTo CustomColor_Err
        '</EhHeader>

        ' If first time, initialize to system colors

100     If fNotFirst = False Then InitColors

104     If i >= 0 And i <= 15 Then

108         alCustom(i) = iValue

        End If

        '<EhFooter>
        Exit Property

CustomColor_Err:
112     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.CustomColor", Erl, False
116     Resume Next
        '</EhFooter>
End Property

' ChooseFont wrapper
Function VBChooseFont(CurFont As Font, _
   Optional PrinterDC As Long = -1, _
   Optional Owner As Long = -1, _
   Optional Color As Long = vbBlack, _
   Optional MinSize As Long = 0, _
   Optional MaxSize As Long = 0, _
   Optional flags As Long = 0) As Boolean
        '<EhHeader>
        On Error GoTo VBChooseFont_Err
        '</EhHeader>

100     m_lApiReturn = 0
104     m_lExtendedError = 0

        ' Unwanted Flags bits
        Const CF_FontNotSupported = CF_Apply Or CF_EnableHook Or CF_EnableTemplate
    
        ' Flags can get reference variable or constant with bit flags
        ' PrinterDC can take printer DC

108     If PrinterDC = -1 Then

112         PrinterDC = 0

116         If flags And CF_PrinterFonts Then PrinterDC = Printer.hdc

        Else

120         flags = flags Or CF_PrinterFonts

        End If

        ' Must have some fonts

124     If (flags And CF_PrinterFonts) = 0 Then flags = flags Or CF_ScreenFonts

        ' Color can take initial color, receive chosen color

128     If Color <> vbBlack Then flags = flags Or CF_EFFECTS

        ' MinSize can be minimum size accepted

132     If MinSize Then flags = flags Or CF_LimitSize

        ' MaxSize can be maximum size accepted

136     If MaxSize Then flags = flags Or CF_LimitSize

        ' Put in required internal flags and remove unsupported
140     flags = (flags Or CF_InitToLogFontStruct) And Not CF_FontNotSupported
    
        ' Initialize LOGFONT variable
        Dim fnt As LOGFONT
        Const PointsPerTwip = 1440 / 72
144     fnt.lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
148     fnt.lfWeight = CurFont.Weight
152     fnt.lfItalic = CurFont.Italic
156     fnt.lfUnderline = CurFont.Underline
160     fnt.lfStrikeOut = CurFont.Strikethrough
        ' Other fields zero
164     StrToBytes fnt.lfFaceName, CurFont.Name

        ' Initialize TCHOOSEFONT variable
        Dim cf As TCHOOSEFONT
168     cf.lStructSize = Len(cf)

172     If Owner <> -1 Then cf.hWndOwner = Owner

176     cf.hdc = PrinterDC
180     cf.lpLogFont = VarPtr(fnt)
184     cf.iPointSize = CurFont.Size * 10
188     cf.flags = flags
192     cf.rgbColors = Color
196     cf.nSizeMin = MinSize
200     cf.nSizeMax = MaxSize
    
        ' All other fields zero
204     m_lApiReturn = ChooseFont(cf)

208     Select Case m_lApiReturn

            Case 1
                ' Success
212             VBChooseFont = True
216             flags = cf.flags
220             Color = cf.rgbColors
224             CurFont.Bold = cf.nFontType And Bold_FontType
                'CurFont.Italic = cf.nFontType And Italic_FontType
228             CurFont.Italic = fnt.lfItalic
232             CurFont.Strikethrough = fnt.lfStrikeOut
236             CurFont.Underline = fnt.lfUnderline
240             CurFont.Weight = fnt.lfWeight
244             CurFont.Size = cf.iPointSize / 10
248             CurFont.Name = BytesToStr(fnt.lfFaceName)

252         Case 0
                ' Cancelled
256             VBChooseFont = False

260         Case Else
                ' Extended error
264             m_lExtendedError = CommDlgExtendedError()
268             VBChooseFont = False

        End Select
        
        '<EhFooter>
        Exit Function

VBChooseFont_Err:
272     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.VBChooseFont", Erl, False
276     Resume Next
        '</EhFooter>
End Function

' PrintDlg wrapper
Function VBPrintDlg(hdc As Long, _
   Optional PrintRange As EPrintRange = eprAll, _
   Optional DisablePageNumbers As Boolean, _
   Optional FromPage As Long = 1, _
   Optional ToPage As Long = &HFFFF, _
   Optional DisableSelection As Boolean, _
   Optional Copies As Integer, _
   Optional ShowPrintToFile As Boolean, _
   Optional DisablePrintToFile As Boolean = True, _
   Optional PrintToFile As Boolean, _
   Optional Collate As Boolean, _
   Optional PreventWarning As Boolean, _
   Optional Owner As Long, _
   Optional Printer As Object, _
   Optional flags As Long) As Boolean
        '<EhHeader>
        On Error GoTo VBPrintDlg_Err
        '</EhHeader>

        Dim afFlags As Long, afMask As Long
    
100     m_lApiReturn = 0
104     m_lExtendedError = 0
    
        ' Set PRINTDLG flags
108     afFlags = (-DisablePageNumbers * PD_NOPAGENUMS) Or _
           (-DisablePrintToFile * PD_DISABLEPRINTTOFILE) Or _
           (-DisableSelection * PD_NOSELECTION) Or _
           (-PrintToFile * PD_PRINTTOFILE) Or _
           (-(Not ShowPrintToFile) * PD_HIDEPRINTTOFILE) Or _
           (-PreventWarning * PD_NOWARNING) Or _
           (-Collate * PD_COLLATE) Or _
           PD_USEDEVMODECOPIESANDCOLLATE Or _
           PD_RETURNDC

112     If PrintRange = eprPageNumbers Then

116         afFlags = afFlags Or PD_PAGENUMS

120     ElseIf PrintRange = eprSelection Then

124         afFlags = afFlags Or PD_SELECTION

        End If

        ' Mask out unwanted bits
128     afMask = CLng(Not (PD_ENABLEPRINTHOOK Or _
           PD_ENABLEPRINTTEMPLATE))
132     afMask = afMask And _
           CLng(Not (PD_ENABLESETUPHOOK Or _
           PD_ENABLESETUPTEMPLATE))
    
        ' Fill in PRINTDLG structure
        Dim pd As TPRINTDLG
136     pd.lStructSize = Len(pd)
140     pd.hWndOwner = Owner
144     pd.flags = afFlags And afMask
148     pd.nFromPage = FromPage
152     pd.nToPage = ToPage
156     pd.nMinPage = 1
160     pd.nMaxPage = &HFFFF
    
        ' Show Print dialog
164     m_lApiReturn = PrintDlg(pd)

168     Select Case m_lApiReturn

            Case 1
172             VBPrintDlg = True
                ' Return dialog values in parameters
176             hdc = pd.hdc

180             If (pd.flags And PD_PAGENUMS) Then

184                 PrintRange = eprPageNumbers

188             ElseIf (pd.flags And PD_SELECTION) Then

192                 PrintRange = eprSelection

                Else

196                 PrintRange = eprAll

                End If

200             FromPage = pd.nFromPage
204             ToPage = pd.nToPage
208             PrintToFile = (pd.flags And PD_PRINTTOFILE)
                ' Get DEVMODE structure from PRINTDLG
                Dim pDevMode As Long
212             pDevMode = GlobalLock(pd.hDevMode)
216             CopyMemory m_dvmode, ByVal pDevMode, Len(m_dvmode)
220             Call GlobalUnlock(pd.hDevMode)
                ' Get Copies and Collate settings from DEVMODE structure
224             Copies = m_dvmode.dmCopies
228             Collate = (m_dvmode.dmCollate = DMCOLLATE_TRUE)
                
                ' Set default printer properties
                On Error Resume Next

232             If Not (Printer Is Nothing) Then

236                 Printer.Copies = Copies
240                 Printer.Orientation = m_dvmode.dmOrientation
244                 Printer.PaperSize = m_dvmode.dmPaperSize
248                 Printer.PrintQuality = m_dvmode.dmPrintQuality

                End If

                On Error GoTo VBPrintDlg_Err

252         Case 0
                ' Cancelled
256             VBPrintDlg = False

260         Case Else
                ' Extended error:
264             m_lExtendedError = CommDlgExtendedError()
268             VBPrintDlg = False

        End Select
    
        '<EhFooter>
        Exit Function

VBPrintDlg_Err:
272     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.VBPrintDlg", Erl, False
276     Resume Next
        '</EhFooter>
End Function

Private Property Get DevMode() As DevMode
        '<EhHeader>
        On Error GoTo DevMode_Err
        '</EhHeader>

100     DevMode = m_dvmode

        '<EhFooter>
        Exit Property

DevMode_Err:
104     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.DevMode", Erl, False
108     Resume Next
        '</EhFooter>
End Property

' PageSetupDlg wrapper
Function VBPageSetupDlg(Optional Owner As Long, _
   Optional DisableMargins As Boolean, _
   Optional DisableOrientation As Boolean, _
   Optional DisablePaper As Boolean, _
   Optional DisablePrinter As Boolean, _
   Optional LeftMargin As Long, _
   Optional MinLeftMargin As Long, _
   Optional RightMargin As Long, _
   Optional MinRightMargin As Long, _
   Optional TopMargin As Long, _
   Optional MinTopMargin As Long, _
   Optional BottomMargin As Long, _
   Optional MinBottomMargin As Long, _
   Optional PaperSize As EPaperSize = epsLetter, _
   Optional Orientation As EOrientation = eoPortrait, _
   Optional PrintQuality As EPrintQuality = epqDraft, _
   Optional Units As EPageSetupUnits = epsuInches, _
   Optional Printer As Object, _
   Optional flags As Long) As Boolean
        '<EhHeader>
        On Error GoTo VBPageSetupDlg_Err
        '</EhHeader>

        Dim afFlags As Long, afMask As Long
        
100     m_lApiReturn = 0
104     m_lExtendedError = 0
        ' Mask out unwanted bits
108     afMask = Not (PSD_EnablePagePaintHook Or _
           PSD_EnablePageSetupHook Or _
           PSD_EnablePageSetupTemplate)
        ' Set TPAGESETUPDLG flags
112     afFlags = (-DisableMargins * PSD_DISABLEMARGINS) Or _
           (-DisableOrientation * PSD_DISABLEORIENTATION) Or _
           (-DisablePaper * PSD_DISABLEPAPER) Or _
           (-DisablePrinter * PSD_DISABLEPRINTER) Or _
           PSD_MARGINS Or PSD_MINMARGINS And afMask
        Dim lUnits As Long

116     If Units = epsuInches Then

120         afFlags = afFlags Or PSD_INTHOUSANDTHSOFINCHES
124         lUnits = 1000

        Else

128         afFlags = afFlags Or PSD_INHUNDREDTHSOFMILLIMETERS
132         lUnits = 100

        End If
    
        Dim psd As TPAGESETUPDLG
        ' Fill in PRINTDLG structure
136     psd.lStructSize = Len(psd)
140     psd.hWndOwner = Owner
144     psd.rtMargin.TOp = TopMargin * lUnits
148     psd.rtMargin.Left = LeftMargin * lUnits
152     psd.rtMargin.Bottom = BottomMargin * lUnits
156     psd.rtMargin.Right = RightMargin * lUnits
160     psd.rtMinMargin.TOp = MinTopMargin * lUnits
164     psd.rtMinMargin.Left = MinLeftMargin * lUnits
168     psd.rtMinMargin.Bottom = MinBottomMargin * lUnits
172     psd.rtMinMargin.Right = MinRightMargin * lUnits
176     psd.flags = afFlags
    
        ' Show Print dialog

180     If PageSetupDlg(psd) Then

184         VBPageSetupDlg = True
            ' Return dialog values in parameters
188         TopMargin = psd.rtMargin.TOp / lUnits
192         LeftMargin = psd.rtMargin.Left / lUnits
196         BottomMargin = psd.rtMargin.Bottom / lUnits
200         RightMargin = psd.rtMargin.Right / lUnits
204         MinTopMargin = psd.rtMinMargin.TOp / lUnits
208         MinLeftMargin = psd.rtMinMargin.Left / lUnits
212         MinBottomMargin = psd.rtMinMargin.Bottom / lUnits
216         MinRightMargin = psd.rtMinMargin.Right / lUnits
        
            ' Get DEVMODE structure from PRINTDLG
            Dim dvmode As DevMode, pDevMode As Long
220         pDevMode = GlobalLock(psd.hDevMode)
224         CopyMemory dvmode, ByVal pDevMode, Len(dvmode)
228         Call GlobalUnlock(psd.hDevMode)
232         PaperSize = dvmode.dmPaperSize
236         Orientation = dvmode.dmOrientation
240         PrintQuality = dvmode.dmPrintQuality
            ' Set default printer properties
            On Error Resume Next

244         If Not (Printer Is Nothing) Then

248             Printer.Copies = dvmode.dmCopies
252             Printer.Orientation = dvmode.dmOrientation
256             Printer.PaperSize = dvmode.dmPaperSize
260             Printer.PrintQuality = dvmode.dmPrintQuality

            End If

            On Error GoTo VBPageSetupDlg_Err

        End If

        '<EhFooter>
        Exit Function

VBPageSetupDlg_Err:
264     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.VBPageSetupDlg", Erl, False
268     Resume Next
        '</EhFooter>
End Function

#If fComponent = 0 Then

    Private Sub ErrRaise(e As Long)
        '<EhHeader>
        On Error GoTo ErrRaise_Err
        '</EhHeader>

            Dim sText As String, sSource As String

100         If e > 1000 Then

104             sSource = App.EXEName & ".CommonDialog"
108             Err.Raise COMError(e), sSource, sText

            Else

                ' Raise standard Visual Basic error
112             sSource = App.EXEName & ".VBError"
116             Err.Raise e, sSource

            End If

        '<EhFooter>
        Exit Sub

ErrRaise_Err:
120     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.ErrRaise", Erl, False
124     Resume Next
        '</EhFooter>
    End Sub


#End If

Private Sub StrToBytes(ab() As Byte, s As String)
        '<EhHeader>
        On Error GoTo StrToBytes_Err
        '</EhHeader>

100     If IsArrayEmpty(ab) Then

            ' Assign to empty array
104         ab = StrConv(s, vbFromUnicode)

        Else

            Dim cab As Long
            ' Copy to existing array, padding or truncating if necessary
108         cab = UBound(ab) - LBound(ab) + 1

112         If Len(s) < cab Then s = s & String$(cab - Len(s), 0)

            'If UnicodeTypeLib Then
            '    Dim st As String
            '    st = StrConv(s, vbFromUnicode)
            '    CopyMemoryStr ab(LBound(ab)), st, cab
            'Else
116         CopyMemoryStr ab(LBound(ab)), s, cab
            'End If

        End If

        '<EhFooter>
        Exit Sub

StrToBytes_Err:
120     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.StrToBytes", Erl, False
124     Resume Next
        '</EhFooter>
End Sub

Private Function BytesToStr(ab() As Byte) As String
        '<EhHeader>
        On Error GoTo BytesToStr_Err
        '</EhHeader>

100     BytesToStr = StrConv(ab, vbUnicode)

        '<EhFooter>
        Exit Function

BytesToStr_Err:
104     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.BytesToStr", Erl, False
108     Resume Next
        '</EhFooter>
End Function

Private Function COMError(e As Long) As Long
        '<EhHeader>
        On Error GoTo COMError_Err
        '</EhHeader>

100     COMError = e Or vbObjectError

        '<EhFooter>
        Exit Function

COMError_Err:
104     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.COMError", Erl, False
108     Resume Next
        '</EhFooter>
End Function

'
Private Function IsArrayEmpty(va As Variant) As Boolean
        '<EhHeader>
        On Error GoTo IsArrayEmpty_Err
        '</EhHeader>

        Dim v As Variant
        On Error Resume Next
100     v = va(LBound(va))
104     IsArrayEmpty = (Err <> 0)

        '<EhFooter>
        Exit Function

IsArrayEmpty_Err:
108     DisplayErrMsg Err.Description, "WinUI.cCommonDialog.IsArrayEmpty", Erl, False
112     Resume Next
        '</EhFooter>
End Function


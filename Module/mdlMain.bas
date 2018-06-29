Attribute VB_Name = "mdlMain"
'**********************************************
'      NotePad Clone Modules : appMain
'----------------------------------------------
'Created     : 10/22/2012
'Re-Modified : 09/06/2014
'**********************************************

Option Explicit
Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
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
    lfFaceName As String * 31
End Type
Private Type CHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
Private Type PRINTDLG_TYPE
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
Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
Private Type DEVMODE_TYPE
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
Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINESCROLL As Long = &HB6
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)
Public There As Boolean
Public ret As Long
Public TXFileName As String
Public PasteKey As Boolean
Public MyMsg As String
Public CurColumn As Long
Public ToolbarOn As Boolean
Public FormatbarOn As Boolean
Public StatusbarOn As Boolean
Public RulerOn As Boolean
Public WrapOn As Integer
Public i As Integer
Public fCancel As Boolean
Public TXMeasurement As Integer
Public LineSpacing As Long
Public ColorHighLite As Long
Public HighColor As Boolean
Public tRow As Integer
Public tCol As Integer
Public tWidth As Long
Public tColWidth(1 To 20)
Public tCenter As Boolean

Private Type CHOOSECOLOR
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
Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const Pic_paste = &H302
Private RTF_b As RichTextBox
Private mHwnd As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type CMDialog
    ownerform As Long
    filefilter As String
    filetitle As String
    filefilterindex As Long
    FileName As String
    initdir As String
    dialogtitle As String
    flags As Long
End Type
Public Type BBfont
    mFontName As String
    mFontsize As Integer
    mBold As Boolean
    mItalic As Boolean
    mUnderline As Boolean
    mStrikethru As Boolean
    mFontColor As Long
End Type
Public SelectFont As BBfont
Public cmndlg As CMDialog
Dim CustomColors() As Byte
Public Fonting As Boolean
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Type FINDREPLACE
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    flags As Long
    lpstrFindWhat As Long
    lpstrReplaceWith As Long
    wFindWhatLen As Integer
    wReplaceWithLen As Integer
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    ptX As Long
    ptY As Long
End Type

Private Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA" (pFindreplace As Long) As Long
Private Declare Function ReplaceText Lib "comdlg32.dll" Alias "ReplaceTextA" (pFindreplace As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Private Declare Function IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As msg) As Long
Private Declare Function CopyPointer2String Lib "kernel32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProcessHeap& Lib "kernel32" ()
Private Declare Function HeapAlloc& Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long)
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long

Private Const GWL_WNDPROC = (-4)
Private Const HEAP_ZERO_MEMORY = &H8
Public Const FR_DIALOGTERM = &H40
Public Const FR_DOWN = &H1
Public Const FR_ENABLEHOOK = &H100
Public Const FR_ENABLETEMPLATE = &H200
Public Const FR_ENABLETEMPLATEHANDLE = &H2000
Public Const FR_FINDNEXT = &H8
Public Const FR_HIDEMATCHCASE = &H8000
Public Const FR_HIDEUPDOWN = &H4000
Public Const FR_HIDEWHOLEWORD = &H10000
Public Const FR_MATCHCASE = &H4
Public Const FR_NOMATCHCASE = &H800
Public Const FR_NOUPDOWN = &H400
Public Const FR_NOWHOLEWORD = &H1000
Public Const FR_REPLACE = &H10
Public Const FR_REPLACEALL = &H20
Public Const FR_SHOWHELP = &H80
Public Const FR_WHOLEWORD = &H2
Const WM_DESTROY = &H2

Const FINDMSGSTRING = "commdlg_FindReplace"
Const HELPMSGSTRING = "commdlg_help"
Const BufLength = 256

Public hDialog As Long, OldProc As Long
Dim uFindMsg As Long, uHelpMsg As Long, lHeap As Long
Public RetFrs As FINDREPLACE, TMsg As msg
Dim arrFind() As Byte, arrReplace() As Byte
Dim objTarget As Object

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (Prop As SHELLEXECUTEINFO) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Const CB_FINDSTRINGEXACT = &H158
Public Const EM_SCROLL = &HB5
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public FileChanged As Boolean
Public ChangeState As Boolean
Public NoStatusUpdate  As Boolean
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)
Private Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0
Public mEXEpath As String
Public Sub InitCmnDlg(mOwner As Long)
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    Dim i As Integer
    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 254
    Next
    cmndlg.ownerform = mOwner
End Sub
Public Function ShowColor() As Long
    Dim cc As CHOOSECOLOR, mcc As Long
    Dim Custcolor(16) As Long
    Dim lReturn As Long
    cc.lStructSize = Len(cc)
    cc.hwndOwner = cmndlg.ownerform
    cc.hInstance = App.hInstance
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 0
    If CHOOSECOLOR(cc) <> 0 Then
        mcc = cc.rgbResult
        If mcc < 0 Then mcc = -mcc
        If mcc > vbWhite Then mcc = vbWhite
        ShowColor = mcc
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function
Public Sub OpenFile()
    Dim OFName As OPENFILENAME
    With cmndlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.filefilter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .initdir
        OFName.lpstrTitle = .dialogtitle
        OFName.nFilterIndex = .filefilterindex
        OFName.flags = .flags
        If GetOpenFileName(OFName) Then
            .FileName = StripTerminator(Trim$(OFName.lpstrFile))
            .filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            .filefilterindex = OFName.nFilterIndex
        End If
    End With
End Sub
Public Sub SaveFile()
    Dim OFName As OPENFILENAME
    With cmndlg
        OFName.lStructSize = Len(OFName)
        OFName.hwndOwner = .ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.filefilter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .initdir
        OFName.lpstrTitle = .dialogtitle
        OFName.nFilterIndex = .filefilterindex
        OFName.flags = .flags
        If GetSaveFileName(OFName) Then
            .FileName = StripTerminator(Trim$(OFName.lpstrFile))
            .filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            .filefilterindex = OFName.nFilterIndex
        End If
    End With
End Sub
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Function ShowFont() As Boolean
    Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim retval As Long
    lfont.lfHeight = 0
    If SelectFont.mBold Then
        lfont.lfWidth = 700
    Else
        lfont.lfWidth = 0
    End If
    lfont.lfItalic = SelectFont.mItalic
    lfont.lfUnderline = SelectFont.mUnderline
    lfont.lfStrikeOut = SelectFont.mStrikethru
    lfont.lfEscapement = 0
    lfont.lfOrientation = 0
    lfont.lfHeight = SelectFont.mFontsize * 1.33
    lfont.lfWeight = FW_NORMAL
    lfont.lfCharSet = DEFAULT_CHARSET
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    lfont.lfQuality = DEFAULT_QUALITY
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    lfont.lfFaceName = SelectFont.mFontName & vbNullChar
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, lfont, Len(lfont)
    cf.lStructSize = Len(cf)
    cf.hwndOwner = cmndlg.ownerform
    cf.lpLogFont = pMem
    cf.iPointSize = SelectFont.mFontsize * 10
    cf.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT
    cf.rgbColors = SelectFont.mFontColor
    cf.nFontType = REGULAR_FONTTYPE
    cf.nSizeMin = 10
    cf.nSizeMax = 72
    retval = CHOOSEFONT(cf)
    If retval <> 0 Then
        ShowFont = True
        CopyMemory lfont, ByVal pMem, Len(lfont)
        With SelectFont
            .mFontName = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
            .mBold = False
            .mItalic = False
            .mUnderline = False
            .mStrikethru = False
            .mFontsize = cf.iPointSize / 10
            If lfont.lfWeight = 700 Then .mBold = True
            .mItalic = lfont.lfItalic
            .mUnderline = lfont.lfUnderline
            .mStrikethru = lfont.lfStrikeOut
            .mFontColor = cf.rgbColors
        End With
    Else
        ShowFont = False
    End If
    retval = GlobalUnlock(hMem)
    retval = GlobalFree(hMem)
End Function
Public Function ShowPageSetupDlg() As Long
    Dim m_PSD As PAGESETUPDLG
    m_PSD.lStructSize = Len(m_PSD)
    m_PSD.hwndOwner = cmndlg.ownerform
    m_PSD.hInstance = App.hInstance
    m_PSD.flags = 0
    If PAGESETUPDLG(m_PSD) Then
        ShowPageSetupDlg = 0
    Else
        ShowPageSetupDlg = -1
    End If
End Function
Public Function ShowPrinter(Optional PrintFlags As Long) As Boolean
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE
    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String
    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = cmndlg.ownerform
    PrintDlg.flags = PrintFlags
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    On Error GoTo 0
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With
    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If
    If PrintDialog(PrintDlg) <> 0 Then
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                End If
            Next
        End If
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        On Error GoTo 0
        ShowPrinter = True
    Else
        ShowPrinter = False
    End If
 frmMain.rtf.SelPrint (Printer.hDC)
End Function

Public Function SpecialFolder(ByVal CSIDL As Long) As String
Dim r As Long
Dim sPath As String
Dim IDL As ITEMIDLIST
Const NOERROR = 0
Const MAX_LENGTH = 260
r = SHGetSpecialFolderLocation(frmMain.hwnd, CSIDL, IDL)
If r = NOERROR Then
    sPath = Space$(MAX_LENGTH)
    r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
    If r Then
        SpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End If
End Function
Public Sub GetPropDlg(frm As Form, mfile As String)
    Dim Prop As SHELLEXECUTEINFO
    Dim r As Long
    With Prop
        .cbSize = Len(Prop)
        .fMask = &HC
        .hwnd = frm.hwnd
        .lpVerb = "properties"
        .lpFile = mfile
    End With
    r = ShellExecuteEX(Prop)
End Sub
Public Function FileExists(sSource As String) As Boolean
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function
Public Sub FileSave(Text As String, filepath As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open filepath For Binary As #f
    Put #f, , Text
    Close #f
    Exit Sub
End Sub
Public Function OneGulp(Src As String) As String
    On Error Resume Next
    Dim f As Integer, Temp As String
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    Temp = String(LOF(f), Chr$(0))
    Get #f, , Temp
    Close #f
    If Left(Temp, 2) = "ÿþ" Or Left(Temp, 2) = "þÿ" Then Temp = Replace(Right(Temp, Len(Temp) - 2), Chr(0), "")
    OneGulp = Temp
End Function
Public Function PathOnly(ByVal filepath As String) As String
    Dim Temp As String
    Temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(Temp, 1) = "\" Then Temp = Left(Temp, Len(Temp) - 1)
    PathOnly = Temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(filepath, InStrRev(filepath, ".") + 1)
    If dot = True Then ExtOnly = "." + ExtOnly
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    Dim Temp As String
    If InStr(1, filepath, ".") = 0 Then
        Temp = filepath
    Else
        Temp = Mid$(filepath, 1, InStrRev(filepath, "."))
        Temp = Left(Temp, Len(Temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = Temp + newext
End Function
Public Function GetFileSize(zLen As Long) As String
    Dim tmp As String
    Const KB As Double = 1024
    Const MB As Double = 1024 * KB
    Const GB As Double = 1024 * MB
    If zLen < KB Then
        tmp = Format(zLen) & " bytes"
    ElseIf zLen < MB Then
        tmp = Format(zLen / KB, "0.00") & " KB"
    Else
        If zLen / MB > 1024 Then
            tmp = Format(zLen / GB, "0.00") & " GB"
        Else
            tmp = Format(zLen / MB, "0.00") & " MB"
        End If
    End If
    GetFileSize = Chr(32) + tmp + Chr(32)
End Function
Public Sub SetScrollPos(mPos As Long, mRTF As RichTextBox)
    Dim CurLineCount As Long, curvl As Long, lastvl As Long
    CurLineCount = SendMessage(mRTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
    curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    If mPos < curvl Then
        Do Until curvl < mPos
            SendMessage mRTF.hwnd, EM_SCROLL, 2, 0
            curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
            If curvl = 0 Or curvl = CurLineCount Then Exit Do
        Loop
    Else
        Do Until curvl > mPos
            SendMessage mRTF.hwnd, EM_SCROLL, 3, 0
            curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
            If curvl = 0 Or curvl = CurLineCount Or lastvl = curvl Then Exit Do
            lastvl = curvl
        Loop
    End If
    Do Until curvl = mPos
        If mPos < curvl Then
            SendMessage mRTF.hwnd, EM_SCROLL, 0, 0
        Else
            SendMessage mRTF.hwnd, EM_SCROLL, 1, 0
        End If
        curvl = SendMessage(mRTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
        If curvl = 0 Or curvl = CurLineCount Or lastvl = curvl Then Exit Do
        lastvl = curvl
Loop
End Sub
Public Function GetLongFilename(ByVal sShortFilename As String) As String
    Dim lRet As Long
    Dim sLongFilename As String
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    If lRet > Len(sLongFilename) Then
        sLongFilename = String$(lRet + 1, " ")
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    If lRet > 0 Then
        GetLongFilename = Left$(sLongFilename, lRet)
    End If
End Function

Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
    Dim lRegResult As Long
    lRegResult = RegDeleteKey(hKey, strPath)
End Sub
Private Function GetAllValues(hKey As Long, strPath As String) As Boolean
    Dim lRegResult As Long
    Dim hCurKey As Long
    Dim lValueNameSize As Long
    Dim strValueName As String
    Dim lCounter As Long
    Dim byDataBuffer(4000) As Byte
    Dim lDataBufferSize As Long
    Dim lValueType As Long
    Dim intZeroPos As Integer
    Dim z As Long
    Dim nColl As Collection
    Set nColl = New Collection
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    Do
        lValueNameSize = 255
        strValueName = String$(lValueNameSize, " ")
        lDataBufferSize = 4000
        lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
        If lRegResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strValueName, Chr$(0))
            If intZeroPos <> 0 Then nColl.Add strValueName
        Else
            Exit Do
        End If
    Loop
    For z = 1 To nColl.count
        If nColl(z) = App.EXEName + ".exe" Then
            GetAllValues = True
            Exit Function
        End If
    Next z
End Function
Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strbuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    If Not IsEmpty(Default) Then
        GetSettingString = Default
    Else
        GetSettingString = ""
    End If
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strbuffer = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strbuffer, lDataBufferSize)
            intZeroPos = InStr(strbuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetSettingString = Left$(strbuffer, intZeroPos - 1)
            Else
                GetSettingString = strbuffer
            End If
        End If
    Else
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function
Private Function GetWinDir() As String
    Dim Path As String, strSave As String
    strSave = String(200, Chr$(0))
    GetWinDir = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave)))
End Function
Private Function SysDir() As String
    Dim sSave As String, ret As Long
    sSave = Space(255)
    ret = GetSystemDirectory(sSave, 255)
    sSave = Left$(sSave, ret)
    SysDir = sSave
End Function
Public Function strUnQuoteString(ByVal strQuotedString As String)
    strQuotedString = Trim$(strQuotedString)
    If Mid$(strQuotedString, 1, 1) = Chr(34) Then
        If Right$(strQuotedString, 1) = Chr(34) Then
            strQuotedString = Mid$(strQuotedString, 2, Len(strQuotedString) - 2)
        End If
    End If
    strUnQuoteString = strQuotedString
End Function
Public Sub AssociateText()
    Dim EXpath As String
    If Right(App.Path, 1) = "\" Then
        EXpath = App.Path + App.EXEName + ".exe %1"
    Else
        EXpath = App.Path + "\" + App.EXEName + ".exe %1"
    End If
    SaveSettingString HKEY_CLASSES_ROOT, ".txt", "", App.EXEName + ".TXT"
    SaveSettingString HKEY_CLASSES_ROOT, ".txt\ShellNew", "", ""
    SaveSettingString HKEY_CLASSES_ROOT, ".txt\ShellNew", "NullFile", ""
    SaveSettingString HKEY_CLASSES_ROOT, App.EXEName + ".TXT", "", "Text Document"
    SaveSettingString HKEY_CLASSES_ROOT, App.EXEName + ".TXT" + "\shell\open\command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, App.EXEName + ".TXT" + "\DefaultIcon", "", SysDir + "\shell32.dll,-152"
      SaveSettingString HKEY_CLASSES_ROOT, ".bpx", "", App.EXEName + ".BPX"
    SaveSettingString HKEY_CLASSES_ROOT, ".bpx\ShellNew", "", ""
    SaveSettingString HKEY_CLASSES_ROOT, ".bpx\ShellNew", "NullFile", ""
    SaveSettingString HKEY_CLASSES_ROOT, App.EXEName + ".bpx" + "\shell\open\command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, App.EXEName + ".BPX" + "\DefaultIcon", "", SysDir + "\shell32.dll,-152"
    SaveSettingString HKEY_CLASSES_ROOT, ".log", "", App.EXEName + ".TXT"
    SaveSettingString HKEY_CLASSES_ROOT, "inifile\shell\open\command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "inffile\shell\open\command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "batfile\shell\edit\command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "JSEFile\Shell\Edit\Command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "JSFile\Shell\Edit\Command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "scpfile\shell\open\command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "VBEFile\Shell\Edit\Command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "VBSFile\Shell\Edit\Command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "WSFFile\Shell\Edit\Command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, "regfile\shell\edit\command", "", EXpath
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
Public Sub AssociateNotepad()
    SaveSettingString HKEY_CLASSES_ROOT, ".txt", "", "txtfile"
    SaveSettingString HKEY_CLASSES_ROOT, ".log", "", "txtfile"
    SaveSettingString HKEY_CLASSES_ROOT, "inifile\shell\open\command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "inffile\shell\open\command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "batfile\shell\edit\command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "JSEFile\Shell\Edit\Command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "JSFile\Shell\Edit\Command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "scpfile\shell\open\command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "VBEFile\Shell\Edit\Command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "VBSFile\Shell\Edit\Command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "WSFFile\Shell\Edit\Command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SaveSettingString HKEY_CLASSES_ROOT, "regfile\shell\edit\command", "", GetWinDir + "\NOTEPAD.EXE %1"
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
Public Sub AssociateBPXText()
    Dim EXpath As String
    If Right(App.Path, 1) = "\" Then
        EXpath = App.Path + App.EXEName + ".exe %1"
    Else
        EXpath = App.Path + "\" + App.EXEName + ".exe %1"
    End If
    SaveSettingString HKEY_CLASSES_ROOT, ".bpx", "", App.EXEName + ".BPX"
    SaveSettingString HKEY_CLASSES_ROOT, ".pnt\ShellNew", "", ""
    SaveSettingString HKEY_CLASSES_ROOT, ".pnt\ShellNew", "NullFile", ""
    SaveSettingString HKEY_CLASSES_ROOT, App.EXEName + ".BPX", "", "RotePad"
    SaveSettingString HKEY_CLASSES_ROOT, App.EXEName + ".BPX" + "\shell\open\command", "", EXpath
    SaveSettingString HKEY_CLASSES_ROOT, App.EXEName + ".BPX" + "\DefaultIcon", "", SysDir + "\shell32.dll,-152"
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
Public Sub AssociateNoBPXText()
    Dim EXpath As String
    If Right(App.Path, 1) = "\" Then
        EXpath = App.Path + App.EXEName + ".exe %1"
    Else
        EXpath = App.Path + "\" + App.EXEName + ".exe %1"
    End If
    SaveSettingString HKEY_CLASSES_ROOT, " ", "", " "
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
Public Function IsAssociatedText() As Boolean
    Dim Answer As String
    Answer = GetSettingString(HKEY_CLASSES_ROOT, ".txt", "")
    If UCase(Answer) = UCase(App.EXEName + ".TXT") Then
        IsAssociatedText = True
    Else
        IsAssociatedText = False
    End If
    Answer = GetSettingString(HKEY_CLASSES_ROOT, ".bpx", "")
    If UCase(Answer) = UCase(App.EXEName + ".bpx") Then
        IsAssociatedText = True
    Else
        IsAssociatedText = False
    End If
End Function
Public Function IsNotePadAssociatedText() As Boolean
    Dim Answer As String
    Answer = GetSettingString(HKEY_CLASSES_ROOT, ".txt", "")
    If UCase(Answer) = UCase("txtfile") Then
        IsNotePadAssociatedText = True
    Else
        IsNotePadAssociatedText = False
    End If
End Function
Public Sub AddSCviewer()
   If IsSCviewer Then Exit Sub
    SaveSettingString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\View Source Editor\Editor Name", "", App.EXEName + ".exe"
End Sub
Public Sub RemoveSCviewer()
    If IsSCviewer Then DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\View Source Editor\Editor Name"
End Sub
Public Function IsSCviewer() As Boolean
    Dim Answer As String
    Answer = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\View Source Editor\Editor Name", "")
    If UCase(Answer) = UCase(App.EXEName + ".exe") Then
        IsSCviewer = True
    Else
        IsSCviewer = False
    End If
End Function
Public Sub AddShortCutSendTo()
    Dim EXpath As String
    Dim sh As Object, sShortcutPath As String
    Dim link As Object
    If Right(App.Path, 1) = "\" Then
        EXpath = App.Path + App.EXEName + ".exe"
    Else
        EXpath = App.Path + "\" + App.EXEName + ".exe"
    End If
    Set sh = CreateObject("WScript.Shell")
    sShortcutPath = SpecialFolder(9) + "\RoteNote.lnk"
    If IsObject(sh) Then
        Set link = sh.CreateShortcut(sShortcutPath)
        If IsObject(link) Then
            link.Description = "An Application to use read, write Plain Text or Else"
            link.IconLocation = EXpath
            link.TargetPath = EXpath
            link.WindowStyle = 0
            link.WorkingDirectory = EXpath
            link.Save
        End If
    End If
End Sub

Public Sub ShowFind(fOwner As Form, objWhere As Object, lFlags As Long, sFind As String, Optional bReplace As Boolean = False, Optional sReplace As String = "")
   If hDialog > 0 Then Exit Sub
   Set objTarget = objWhere
   Dim FRS As FINDREPLACE, i As Integer
   arrFind = StrConv(sFind & Chr$(0), vbFromUnicode)
   arrReplace = StrConv(sReplace & Chr$(0), vbFromUnicode)
   With FRS
        .lStructSize = LenB(FRS) '&H20     '
        .lpstrFindWhat = VarPtr(arrFind(0))
        .wFindWhatLen = BufLength
        .lpstrReplaceWith = VarPtr(arrReplace(0))
        .wReplaceWithLen = BufLength
        .hwndOwner = fOwner.hwnd
        .flags = lFlags
        .hInstance = App.hInstance
    End With
    lHeap = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, FRS.lStructSize)
    CopyMemory ByVal lHeap, FRS, Len(FRS)
    uFindMsg = RegisterWindowMessage(FINDMSGSTRING)
    uHelpMsg = RegisterWindowMessage(HELPMSGSTRING)
    OldProc = SetWindowLong(fOwner.hwnd, GWL_WNDPROC, AddressOf WndProc)
    If bReplace Then
       hDialog = ReplaceText(ByVal lHeap)
    Else
       hDialog = FindText(ByVal lHeap)
    End If
    MessageLoop
End Sub

Private Sub MessageLoop()
  Do While GetMessage(TMsg, 0&, 0&, 0&) And hDialog > 0
     If IsDialogMessage(hDialog, TMsg) = False Then
        TranslateMessage TMsg
        DispatchMessage TMsg
     End If
  Loop
End Sub

Public Function WndProc(ByVal hOwner As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case wMsg
      Case uFindMsg
           CopyMemory RetFrs, ByVal lParam, Len(RetFrs)
           If (RetFrs.flags And FR_DIALOGTERM) = FR_DIALOGTERM Then
              SetWindowLong hOwner, GWL_WNDPROC, OldProc
              HeapFree GetProcessHeap(), 0, lHeap
              hDialog = 0: lHeap = 0: OldProc = 0
              If objTarget.HideSelection Then objTarget.SetFocus
              Set objTarget = Nothing
           Else
              DoFindReplace RetFrs
           End If
      Case uHelpMsg
           MsgBox "Here is your code to call your help file", vbInformation + vbOKOnly, "Heeeelp!!!!"
      Case Else
           If wMsg = WM_DESTROY Then
              EndDialog hDialog, 0&
              SetWindowLong hOwner, GWL_WNDPROC, OldProc
              HeapFree GetProcessHeap(), 0, lHeap
              hDialog = 0: lHeap = 0: OldProc = 0
              Set objTarget = Nothing
              Exit Function
           End If
           WndProc = CallWindowProc(OldProc, hOwner, wMsg, wParam, lParam)
   End Select
End Function

Private Sub DoFindReplace(fr As FINDREPLACE)
  If CheckFlags(FR_FINDNEXT, fr.flags) Then
     If CheckFlags(FR_DOWN, fr.flags) Then
        FindNextWord PointerToString(fr.lpstrFindWhat), fr.flags
     Else
        FindPrevWord PointerToString(fr.lpstrFindWhat), fr.flags
     End If
     If objTarget.HideSelection Then objTarget.SetFocus
  End If
  If CheckFlags(FR_REPLACE, fr.flags) Then ReplaceWord PointerToString(fr.lpstrFindWhat), PointerToString(fr.lpstrReplaceWith), fr.flags
  If CheckFlags(FR_REPLACEALL, fr.flags) Then ReplaceAll PointerToString(fr.lpstrFindWhat), PointerToString(fr.lpstrReplaceWith), fr.flags
End Sub

Private Function PointerToString(p As Long) As String
   Dim s As String
   s = String(BufLength, Chr$(0))
   CopyPointer2String s, p
   PointerToString = Left(s, InStr(s, Chr$(0)) - 1)
End Function

Private Function CheckFlags(flag As Long, flags As Long) As Boolean
   CheckFlags = ((flags And flag) = flag)
End Function

Function FindNextWord(sFind As String, lFlags As Long, Optional bShowMsg As Boolean = True) As Boolean
  Dim lStart As Long, pl As String, nl As String
   With objTarget
      lStart = .SelStart + 1
      If .SelLength > 0 Then lStart = lStart + 1
      Do
        lStart = InStr(lStart, .Text, sFind, IIf(CheckFlags(FR_MATCHCASE, lFlags), vbBinaryCompare, vbTextCompare))
        If lStart = 0 Then Exit Do
        If CheckFlags(FR_WHOLEWORD, lFlags) Then
           If lStart = 1 Then pl = " " Else pl = Mid$(.Text, lStart - 1, 1)
           If lStart + Len(sFind) = Len(.Text) Then nl = " " Else nl = Mid$(.Text, lStart + Len(sFind), 1)
           If ValidateWholeWord(pl, nl) Then Exit Do Else lStart = lStart + 1
        Else
           Exit Do
        End If
      Loop
      If lStart > 0 Then
         .SelStart = lStart - 1
         .SelLength = Len(sFind)
         FindNextWord = True
      Else
         FindNextWord = False
         If bShowMsg Then MsgBox "No matches found", vbExclamation, "Find/Replace"
      End If
   End With
End Function

Function FindPrevWord(sFind As String, lFlags As Long) As Boolean
  Dim lStart As Long, pl As String, nl As String
   With objTarget
      lStart = .SelStart - 1
      If lStart < 0 Then lStart = 0
      Do
        lStart = InStrR(lStart, .Text, sFind, IIf(CheckFlags(FR_MATCHCASE, lFlags), vbBinaryCompare, vbTextCompare))
        If lStart <= 0 Then Exit Do
        If CheckFlags(FR_WHOLEWORD, lFlags) Then
           If lStart = 1 Then pl = " " Else pl = Mid$(.Text, lStart - 1, 1)
           If lStart + Len(sFind) = Len(.Text) Then nl = " " Else nl = Mid$(.Text, lStart + Len(sFind), 1)
           If ValidateWholeWord(pl, nl) Then Exit Do Else lStart = lStart - 1
        Else
           Exit Do
        End If
      Loop
      If lStart > 0 Then
         .SelStart = lStart - 1
         .SelLength = Len(sFind)
         FindPrevWord = True
      Else
         FindPrevWord = False
         MsgBox "No matches found", vbExclamation, "Find/Replace"
      End If
   End With
End Function

Function ReplaceWord(sFind As String, sReplace As String, lFlags As Long)
  With objTarget
      If .SelText <> sFind Then
         FindNextWord sFind, lFlags
      Else
         .SelText = sReplace
         FindNextWord sFind, lFlags
      End If
  End With
End Function

Function ReplaceAll(sFind As String, sReplace As String, lFlags As Long)
  Dim nCount As Long
  With objTarget
      .SelStart = 0
      Do
         If FindNextWord(sFind, lFlags, False) Then
            .SelText = sReplace
            nCount = nCount + 1
         Else
            Exit Do
         End If
      Loop
      If nCount > 0 Then
         MsgBox "Text has been searched. " & nCount & " replacements were made.", vbInformation, "Find/Replace"
      Else
         MsgBox "No matches found", vbExclamation, "Find/Replace"
      End If
  End With
End Function

Private Function ValidateWholeWord(PrevLetter As String, NextLetter As String) As Boolean
   Dim sLetters As String
   ValidateWholeWord = True
   sLetters = "abcdefghijklmnoprqstuvwxyz1234567890"
   If InStr(1, sLetters, PrevLetter, vbTextCompare) Or InStr(1, sLetters, NextLetter, vbTextCompare) Then ValidateWholeWord = False
End Function

Private Function InStrR(Optional lStart As Long, Optional sTarget As String, Optional sFind As String, Optional iCompare As Integer) As Long
    Dim cFind As Long, i As Long
    cFind = Len(sFind)
    For i = lStart - cFind + 1 To 1 Step -1
        If StrComp(Mid$(sTarget, i, cFind), sFind, iCompare) = 0 Then
            InStrR = i
            Exit Function
        End If
    Next
End Function

Public Function GetColPos(tBox As Object) As Long
  GetColPos = tBox.SelStart - SendMessageByNum(tBox.hwnd, EM_LINEINDEX, -1&, 0&)
End Function
Public Function GetLineNum(tBox As Object) As Long
  GetLineNum = SendMessageByNum(tBox.hwnd, EM_LINEFROMCHAR, tBox.SelStart, 0&)
End Function
Public Function LastPart(Text As String) As String
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  LastPart = Temp
 Else
  LastPart = Mid$(Temp, i + 1)
 End If
End Function
Public Function FirstPart(Text As String) As String
 Dim Temp As String
 Dim i As Integer
 Temp = Trim$(Text)
 For i = Len(Temp) To 1 Step -1
  If Mid$(Temp, i, 1) = "\" Then Exit For
 Next i
 If i = 0 Then
  FirstPart = Temp
 Else
  FirstPart = Left$(Temp, i - 1)
 End If
End Function

Public Function FileExistsX(FileName As String) As Boolean
On Error GoTo handle
    If FileLen(FileName) >= 0 Then: FileExistsX = True: Exit Function
handle:
    FileExistsX = False
End Function
Public Function WordCount(Text As String) As Long
    Dim dest() As Byte
    Dim i As Long
    If LenB(Text) Then
        ReDim dest(LenB(Text))
        CopyMemory dest(0), ByVal StrPtr(Text), LenB(Text) - 1
        For i = 0 To UBound(dest) Step 2
            If dest(i) > 32 Then
                Do Until dest(i) < 33
                    i = i + 2
                Loop
                WordCount = WordCount + 1
            End If
        Next i
        Erase dest
    Else
        WordCount = 0
    End If
End Function

Public Function FiXX(FileName As String) As Boolean
On Error GoTo handle
    If FileLen(FileName) >= 0 Then: FiXX = True: Exit Function
handle:
    FiXX = False
End Function

Sub Main()
frmMain.Show
End Sub

Public Function LineCount() As Long

    mHwnd = frmMain.rtf.hwnd
    LineCount = SendMessage(mHwnd, EM_GETLINECOUNT, 0&, 0&)
    
End Function

Public Function GetCharFromLine(LineIndex As Long)

    mHwnd = frmMain.rtf.hwnd
    If LineIndex < LineCount Then
      GetCharFromLine = SendMessage(mHwnd, EM_LINEINDEX, LineIndex, 0&)
    End If
    
End Function





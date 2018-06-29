Attribute VB_Name = "mdlCmnDlg"

'Standard Commondialog - nothing special here
'this code is available EVERYWHERE
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
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
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
 Jx.rtf.SelPrint (Printer.hDC)
End Function


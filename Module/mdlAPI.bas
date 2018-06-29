Attribute VB_Name = "mdlVBAPI"
'**********************************************
'      NotePad Clone Modules : API
'----------------------------------------------
'Created     : 10/22/2012
'Re-Modified : 09/06/2014
'**********************************************

Option Explicit

Public Type POINTAPI
    mx As Long
    my As Long
End Type

Public Pnt As POINTAPI

Public Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

Type CharRange
    cpMin As Long
    cpMax As Long
End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_USER = &H400
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_EXSETSEL = (WM_USER + 55)

Public Const EM_LINEINDEX = &HBB
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_LINESCROLL = &HB6


' Win 32 Declarations for View Mode
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWYSIWYG = 2
End Enum

'----------------------------------------------------------------
'
'              Show fileproperties
'
'----------------------------------------------------------------
' FileName (string) is the full path and name of the file
' MyForm   (form)   is the form on wich you call the properties
' return   (long)   if return > 32 no error occured
'----------------------------------------------------------------

Public Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Type SHELLEXECUTEINFO
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
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long

' Move Window is for SetDropHeight
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long


' Used to determine if a program/process is running
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const MAX_PATH& = 260

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type

'------------------------------------------
'Following is for highlight color
Public Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const SCF_SELECTION = &H1&
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const CFM_BACKCOLOR = &H4000000
Public Const LF_FACESIZE = 32
Public Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lLCID As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bReserved1 As Byte
End Type

Dim udtCharFormat As CHARFORMAT2
Private m_SelHColor As OLE_COLOR

'Used for line spacing
Public Const MAX_TAB_STOPS = 32&
Public Const EM_SETPARAFORMAT = &H447
Public Const PFM_LINESPACING = &H100&

Public Type PARAFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    wNumbering As Integer
    wReserved As Integer
    dxStartIndent As Long
    dxRightIndent As Long
    dxOffset As Long
    wAlignment As Integer
    cTabCount As Integer
    lTabStops(0 To MAX_TAB_STOPS - 1) As Long
    dySpaceBefore As Long          ' Vertical spacing before para
    dySpaceAfter As Long           ' Vertical spacing after para
    dyLineSpacing As Long          ' Line spacing depending on Rule
    sStyle As Integer              ' Style handle
    bLineSpacingRule As Byte       ' Rule for line spacing
    bCRC As Byte                   ' Reserved for CRC for rapid searching
    wShadingWeight As Integer      ' Shading in hundredths of a per cent
    wShadingStyle As Integer       ' Nibble 0: style, 1: cfpat, 2: cbpat
    wNumberingStart As Integer     ' Starting value for numbering
    wNumberingStyle As Integer     ' Alignment, roman/arabic, (), ), .,     etc.
    wNumberingTab As Integer       ' Space between 1st indent and 1st-line text
    wBorderSpace As Integer        ' Space between border and text(twips)
    wBorderWidth As Integer        ' Border pen width (twips)
    wBorders As Integer            ' Byte 0: bits specify which borders; Nibble 2: border style; 3: color                                     index*/
End Type


Public Function SelLineSpacing(rtbTarget As RichTextBox, SpacingRule As Long, Optional LineSpacing As Long = 20)
    ' SpacingRule
    ' Type of line spacing. To use this member, set the PFM_SPACEAFTER flag in the dwMask member. This member can be one of the following values.
    ' 0 - Single spacing. The dyLineSpacing member is ignored.
    ' 1 - One-and-a-half spacing. The dyLineSpacing member is ignored.
    ' 2 - Double spacing. The dyLineSpacing member is ignored.
    ' 3 - The dyLineSpacing member specifies the spacingfrom one line to the next, in twips. However, if dyLineSpacing specifies a value that is less than single spacing, the control displays single-spaced text.
    ' 4 - The dyLineSpacing member specifies the spacing from one line to the next, in twips. The control uses the exact spacing specified, even if dyLineSpacing specifies a value that is less than single spacing.
    ' 5 - The value of dyLineSpacing / 20 is the spacing, in lines, from one line to the next. Thus, setting dyLineSpacing to 20 produces single-spaced text, 40 is double spaced, 60 is triple spaced, and so on.

    Dim Para As PARAFORMAT2
    With Para
        .cbSize = Len(Para)
        .dwMask = PFM_LINESPACING
        .bLineSpacingRule = SpacingRule
        .dyLineSpacing = LineSpacing
    End With
    
    SendMessage rtbTarget.hwnd, EM_SETPARAFORMAT, ByVal 0&, Para
End Function
Public Sub HighliteText(NewColor As OLE_COLOR)
        m_SelHColor = NewColor
        udtCharFormat.dwMask = CFM_BACKCOLOR
        udtCharFormat.cbSize = LenB(udtCharFormat)
        udtCharFormat.crBackColor = m_SelHColor
    Call SendMessageByVal(frmMain.rtf.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(udtCharFormat))
End Sub


Public Function IsAppRunning(FEXEName As String)
 Dim uProcess As PROCESSENTRY32
 Dim rProcessFound As Long
 Dim hSnapshot As Long
 Dim szExename As String
 Dim i As Integer

 Const PROCESS_ALL_ACCESS = 0
 Const TH32CS_SNAPPROCESS As Long = 2&

 uProcess.dwSize = Len(uProcess)
 hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
 rProcessFound = ProcessFirst(hSnapshot, uProcess)

 Do While rProcessFound
    i = InStr(1, uProcess.szexeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
    If szExename = LCase(FEXEName) Then
     IsAppRunning = True
     Call CloseHandle(hSnapshot)
     Exit Function
    End If
    
    rProcessFound = ProcessNext(hSnapshot, uProcess)
 Loop

 'not found
 IsAppRunning = False
 Call CloseHandle(hSnapshot)

End Function

Public Function ShowFileProp(ByVal FileName As String, aForm As Form) As Long

'if return <=32 error occured
Dim SEI As SHELLEXECUTEINFO
Dim r As Long
If FileName = "" Then
    ShowFileProp = 0
    Exit Function
    End If
With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hwnd = aForm.hwnd
    .lpVerb = "properties"
    .lpFile = FileName
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
End With
r = ShellExecuteEX(SEI)
ShowFileProp = SEI.hInstApp
End Function



Attribute VB_Name = "mdlCommandNow"
'**********************************************
'        ROTEJx MODULE TEXT COMMAND
'----------------------------------------------
'Modified by : Teztaz Enterprises
'Created     : 10/22/2012
'**********************************************

Option Explicit
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

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
'P 'ublic Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Const CB_FINDSTRINGEXACT = &H158
Public Const EM_SCROLL = &HB5
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public FileChanged As Boolean
Public ChangeState As Boolean
Public NoStatusUpdate  As Boolean
Public Function SpecialFolder(ByVal CSIDL As Long) As String
Dim r As Long
Dim sPath As String
Dim IDL As ITEMIDLIST
Const NOERROR = 0
Const MAX_LENGTH = 260
r = SHGetSpecialFolderLocation(Jx.hwnd, CSIDL, IDL)
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




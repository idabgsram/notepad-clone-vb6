Attribute VB_Name = "mdlFileHandling"
'**********************************************
'           ROTENOTE MODULE BINARYDC
'----------------------------------------------
'Modified by : Teztaz Enterprises
'Created     : 10/22/2012
'**********************************************

Option Explicit
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



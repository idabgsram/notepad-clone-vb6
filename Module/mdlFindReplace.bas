Attribute VB_Name = "mdlFindReplace"
'************************************************
'* API to display Find and find Replace Dialogs *
'************************************************

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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
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


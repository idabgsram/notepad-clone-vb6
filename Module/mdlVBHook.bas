Attribute VB_Name = "mdlUXHook"
'**********************************************
'      NotePad Clone Modules : UxHook
'----------------------------------------------
'Created     : 10/22/2012
'Re-Modified : 09/06/2014
'**********************************************

Option Explicit

Private Const WH_CBT As Long = &H5
Private Const HCBT_ACTIVATE As Long = &H5
Private Const STM_SETICON As Long = &H170
Private Const MODAL_WINDOW_CLASSNAME As String = "#32770"
Private Const SS_ICON As Long = &H3
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const STM_SETIMAGE As Long = &H172
Private Const IMAGE_CURSOR As Long = &H2

Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Boolean
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long


Private pHook As Long
Private phIcon As Long
Private pAniIcon As String

Public Function XMsgBox(ByVal message As String, _
               Optional ByVal MBoxStyle As VbMsgBoxStyle = vbOKOnly, _
               Optional ByVal Title As String = "", _
               Optional ByVal hIcon As Long = 0&) As VbMsgBoxResult
   
   pHook = SetWindowsHookEx(WH_CBT, _
          AddressOf MsgBoxHookProc, _
                     App.hInstance, _
                 GetCurrentThreadId())
                 
   phIcon = hIcon
   
  
   If phIcon <> 0 Then
      MBoxStyle = MBoxStyle And Not (vbCritical)
      MBoxStyle = MBoxStyle And Not (vbExclamation)
      MBoxStyle = MBoxStyle And Not (vbQuestion)
      MBoxStyle = MBoxStyle Or vbInformation
   End If
   
   XMsgBox = MsgBox(message, MBoxStyle, Title)
End Function

Private Function MsgBoxHookProc(ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim ClassNameSize As Long
   Dim sClassName As String
   Dim hIconWnd As Long

   MsgBoxHookProc = CallNextHookEx(pHook, CodeNo, wParam, lParam)
   If CodeNo = HCBT_ACTIVATE Then
      sClassName = Space$(32)
      ClassNameSize = GetClassName(wParam, sClassName, 32)
      If Left$(sClassName, ClassNameSize) <> MODAL_WINDOW_CLASSNAME Then Exit Function
      If phIcon <> 0 Or Len(pAniIcon) <> 0 Then hIconWnd = FindWindowEx(wParam, 0&, "Static", vbNullString)
      Dim oldText, newText, i As Integer
      oldText = Split("Yes;&Yes;No;&No;Cancel;&Cancel;Abort;&Abort;Ignore;&Ignore;Retry;&Retry", ";")
      newText = Split("Save;&Save;Don't Save;&Don't Save;Cancel;&Cancel;Gugur;&Gugur;Abaikan;&Abaikan;Ulang;&Ulang", ";")
      For i = 0 To UBound(oldText)
      hIconWnd = FindWindowEx(wParam, 0&, "Button", CStr(oldText(i)))
                 SetWindowText hIconWnd, CStr(newText(i))
      Next i
      hIconWnd = FindWindowEx(wParam, 0&, "Static", vbNullString)
      If phIcon <> 0 Then SendMessage hIconWnd, STM_SETICON, phIcon, ByVal 0&
      UnhookWindowsHookEx pHook
   End If
End Function




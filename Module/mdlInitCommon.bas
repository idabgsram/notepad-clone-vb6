Attribute VB_Name = "mdlInitCommon"
'**********************************************
'      ROTEJx MODULE COMMON CONTROLS
'----------------------------------------------
'Modified by : Teztaz Enterprises
'Created     : 10/22/2012
'**********************************************

Option Explicit
Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const Pic_paste = &H302
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private RTF_b As RichTextBox
Private mHwnd As Long
Private Type tagInitCommonControlsEx
  lngSize As Long
  lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200
Public Function WindowsInitCom() As Boolean
On Error Resume Next
Dim iccex As tagInitCommonControlsEx
With iccex
  .lngSize = Len(iccex)
  .lngICC = ICC_USEREX_CLASSES
End With
InitCommonControlsEx iccex
WindowsInitCom = CBool(Err = 0)
End Function

Public Function FiXX(filename As String) As Boolean
On Error GoTo handle
    If FileLen(filename) >= 0 Then: FiXX = True: Exit Function
handle:
    FiXX = False
End Function

Sub main()
Jx.Show
End Sub

Public Function LineCount() As Long

    mHwnd = Jx.rtf.hwnd
    LineCount = SendMessage(mHwnd, EM_GETLINECOUNT, 0&, 0&)
    
End Function

Public Function GetCharFromLine(LineIndex As Long)

    mHwnd = Jx.rtf.hwnd
    If LineIndex < LineCount Then
      GetCharFromLine = SendMessage(mHwnd, EM_LINEINDEX, LineIndex, 0&)
    End If
    
End Function



